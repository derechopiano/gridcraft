//! Workbook-owned Python modules. Pure data: deserializing a project never executes it.
use std::collections::BTreeMap;
use std::sync::Arc;

use serde::{Deserialize, Serialize};

pub const MAX_MODULES: usize = 128;
pub const MAX_MODULE_BYTES: usize = 1_048_576;
pub const MAX_PROJECT_BYTES: usize = 4_194_304;
pub const MAX_METADATA_BYTES: usize = 8192;

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct PythonMacroProject {
    /// PEP 723 comment block only. Executable text is never accepted as metadata.
    pub metadata: String,
    pub modules: BTreeMap<String, Arc<str>>,
}

impl<'de> Deserialize<'de> for PythonMacroProject {
    fn deserialize<D: serde::Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        #[derive(Deserialize)]
        #[serde(deny_unknown_fields)]
        struct Stored {
            metadata: String,
            modules: BTreeMap<String, Arc<str>>,
        }
        let stored = Stored::deserialize(deserializer)?;
        let project = Self { metadata: stored.metadata, modules: stored.modules };
        project.validate().map_err(serde::de::Error::custom)?;
        Ok(project)
    }
}

impl Default for PythonMacroProject {
    fn default() -> Self {
        Self { metadata: "# /// script\n# requires-python = \">=3.11\"\n# dependencies = []\n# ///\n".into(), modules: BTreeMap::new() }
    }
}

impl PythonMacroProject {
    pub fn validate(&self) -> Result<(), String> {
        validate_metadata(&self.metadata)?;
        if self.modules.len() > MAX_MODULES {
            return Err("A Python project supports at most 128 modules".into());
        }
        let mut bytes = self.metadata.len();
        let mut names = std::collections::BTreeSet::new();
        for (name, code) in &self.modules {
            validate_module_name(name)?;
            if !names.insert(name.to_ascii_lowercase()) {
                return Err("Python module names must be distinct ignoring case".into());
            }
            if code.len() > MAX_MODULE_BYTES || code.contains('\0') {
                return Err(format!("Python module {name} exceeds 1 MiB or contains a NUL character"));
            }
            bytes = bytes.saturating_add(name.len()).saturating_add(code.len());
            if bytes > MAX_PROJECT_BYTES {
                return Err("Python project source exceeds 4 MiB".into());
            }
        }
        Ok(())
    }
}

pub fn validate_module_name(name: &str) -> Result<(), String> {
    let mut chars = name.chars();
    let first = chars.next().is_some_and(|c| c.is_ascii_alphabetic() || c == '_');
    let reserved = matches!(
        name.to_ascii_lowercase().as_str(),
        "gridcraft"
            | "__main__"
            | "_gridcraft_entry"
            | "con"
            | "prn"
            | "aux"
            | "nul"
            | "com1"
            | "com2"
            | "com3"
            | "com4"
            | "com5"
            | "com6"
            | "com7"
            | "com8"
            | "com9"
            | "lpt1"
            | "lpt2"
            | "lpt3"
            | "lpt4"
            | "lpt5"
            | "lpt6"
            | "lpt7"
            | "lpt8"
            | "lpt9"
    );
    let keyword = matches!(
        name,
        "False"
            | "None"
            | "True"
            | "and"
            | "as"
            | "assert"
            | "async"
            | "await"
            | "break"
            | "class"
            | "continue"
            | "def"
            | "del"
            | "elif"
            | "else"
            | "except"
            | "finally"
            | "for"
            | "from"
            | "global"
            | "if"
            | "import"
            | "in"
            | "is"
            | "lambda"
            | "nonlocal"
            | "not"
            | "or"
            | "pass"
            | "raise"
            | "return"
            | "try"
            | "while"
            | "with"
            | "yield"
    );
    if name.len() > 64 || !first || !chars.all(|c| c.is_ascii_alphanumeric() || c == '_') || reserved || keyword {
        return Err("Use a non-reserved Python module identifier (ASCII letters, digits and underscores, at most 64 characters)".into());
    }
    Ok(())
}

pub fn validate_metadata(metadata: &str) -> Result<(), String> {
    if metadata.len() > MAX_METADATA_BYTES || metadata.contains('\0') {
        return Err("Python metadata exceeds 8 KiB or contains a NUL character".into());
    }
    let lines: Vec<_> = metadata.trim().lines().collect();
    if lines.first().copied() != Some("# /// script") || lines.last().copied() != Some("# ///") || lines.len() < 3 {
        return Err("Provide one PEP 723 script metadata comment block".into());
    }
    if lines
        .iter()
        .skip(1)
        .take(lines.len().saturating_sub(2))
        .any(|line| !(line.is_empty() || *line == "#" || line.starts_with("# ")) || line.starts_with("# ///"))
    {
        return Err("Python project metadata must contain only comments inside one script block".into());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn names_are_portable_and_cannot_escape_the_project() {
        for name in ["", "../escape", "a.py", "C:\\evil", "gridcraft", "GridCraft", "def", "NUL", "com1", "has space", "2bad"] {
            assert!(validate_module_name(name).is_err(), "{name}");
        }
        for name in ["analysis", "stats2", "_helper"] {
            validate_module_name(name).unwrap();
        }
    }
    #[test]
    fn projects_are_bounded_and_names_are_unambiguous() {
        let mut project = PythonMacroProject::default();
        project.modules.insert("analysis".into(), "def run():\n    pass\n".into());
        project.validate().unwrap();
        let json = serde_json::to_string(&project).unwrap();
        assert_eq!(serde_json::from_str::<PythonMacroProject>(&json).unwrap(), project);
        project.modules.insert("Analysis".into(), "".into());
        assert!(project.validate().is_err());
        project.modules.remove("Analysis");
        project.modules.insert("stats".into(), "x".repeat(MAX_MODULE_BYTES + 1).into());
        assert!(project.validate().is_err());
    }
    #[test]
    fn metadata_cannot_inject_executable_source_or_extra_blocks() {
        validate_metadata(&PythonMacroProject::default().metadata).unwrap();
        for metadata in ["print('bad')", "# /// script\nprint('bad')\n# ///", "# /// script\n# ///\n# /// script\n# ///"] {
            assert!(validate_metadata(metadata).is_err());
        }
    }
}
