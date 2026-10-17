//! GridCraft Python macro projects are an inert, versioned custom workbook part.
use std::sync::Arc;

use gridcraft_model::{PythonMacroProject, Workbook};
use serde::{Deserialize, Serialize};

use crate::IoError;

pub(crate) const REL: &str = "urn:gridcraft:relationships/pythonMacroProject";
pub(crate) const CONTENT_TYPE: &str = "application/vnd.gridcraft.python-macro-project+json";
pub(crate) const MAX_PART_BYTES: u64 = 32 * 1024 * 1024;

#[derive(Serialize, Deserialize)]
#[serde(deny_unknown_fields)]
struct Part {
    version: u32,
    project: Arc<PythonMacroProject>,
}

pub(crate) fn read(bytes: &[u8]) -> Result<Arc<PythonMacroProject>, IoError> {
    if bytes.len() as u64 > MAX_PART_BYTES {
        return Err(IoError::TooLarge("Python macro project part exceeds 32 MiB".into()));
    }
    let part: Part = serde_json::from_slice(bytes).map_err(|e| IoError::Format(format!("invalid Python macro project: {e}")))?;
    if part.version != 1 {
        return Err(IoError::Format(format!("unsupported Python macro project version {}", part.version)));
    }
    part.project.validate().map_err(IoError::Format)?;
    Ok(part.project)
}

pub(crate) fn write(wb: &Workbook) -> Result<Option<Vec<u8>>, IoError> {
    let Some(project) = &wb.python_macros else { return Ok(None) };
    project.validate().map_err(IoError::Format)?;
    let bytes = serde_json::to_vec(&Part { version: 1, project: project.clone() }).map_err(|e| IoError::Format(e.to_string()))?;
    if bytes.len() as u64 > MAX_PART_BYTES {
        return Err(IoError::TooLarge("Python macro project part exceeds 32 MiB".into()));
    }
    Ok(Some(bytes))
}

#[cfg(test)]
mod tests {
    use super::*;
    use gridcraft_core::{CellRef, Value};

    #[test]
    fn inert_project_round_trips_beside_cells_and_vba() {
        let mut wb = Workbook::new();
        let mut project = PythonMacroProject::default();
        project.modules.insert("analysis".into(), "raise Exception('opening must never run this')\n".into());
        project.modules.insert("stats".into(), "def mean(values):\n    return sum(values) / len(values)\n".into());
        wb.python_macros = Some(Arc::new(project));
        wb.vba_project = Some(vec![1, 2, 3]);
        wb.sheet_mut(0).unwrap().set_value(CellRef::new(0, 0), Value::Number(7.0));
        let bytes = crate::write_xlsx(&wb).unwrap();
        let (reopened, _) = crate::read_xlsx(&bytes).unwrap();
        assert_eq!(reopened.python_macros, wb.python_macros);
        assert_eq!(reopened.vba_project, wb.vba_project);
        assert_eq!(reopened.sheet(0).unwrap().value(CellRef::new(0, 0)), Value::Number(7.0));
        let again = crate::write_xlsx(&reopened).unwrap();
        assert_eq!(crate::read_xlsx(&again).unwrap().0.python_macros, wb.python_macros);
    }

    #[test]
    fn future_or_malformed_projects_are_not_silently_dropped() {
        assert!(read(br#"{"version":2,"project":{"metadata":"","modules":{}}}"#).is_err());
        assert!(read(br#"{"version":1,"project":{"metadata":"print('bad')","modules":{}}}"#).is_err());
        assert!(read(b"not json").is_err());
        let wb = Workbook::new();
        assert!(write(&wb).unwrap().is_none());
    }
}
