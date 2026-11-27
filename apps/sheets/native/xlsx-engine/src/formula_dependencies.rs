use super::{DefinedName, HashMap, HashSet, SheetMetadata};
use ironcalc::base::{
    expressions::{
        lexer::{Lexer, LexerMode},
        token::TokenType,
    },
    language::get_language,
    locale::get_locale,
};

pub(super) struct DependencyCatalog {
    sheets: HashMap<String, usize>,
    tables: HashMap<String, usize>,
    names: HashMap<(Option<usize>, String), String>,
}

impl DependencyCatalog {
    pub(super) fn new(sheets: &[SheetMetadata], names: &[DefinedName]) -> Self {
        Self {
            sheets: sheets
                .iter()
                .enumerate()
                .map(|(index, sheet)| (sheet.name.to_lowercase(), index))
                .collect(),
            tables: sheets
                .iter()
                .enumerate()
                .flat_map(|(index, sheet)| {
                    sheet.tables.iter().filter_map(move |table| {
                        table.name.as_ref().map(|name| (name.to_lowercase(), index))
                    })
                })
                .collect(),
            names: names
                .iter()
                .map(|name| {
                    (
                        (name.sheet_index, name.name.to_lowercase()),
                        name.formula.clone(),
                    )
                })
                .collect(),
        }
    }

    pub(super) fn scan(
        &self,
        source: usize,
        formula: &str,
        known: &mut HashSet<usize>,
        names_seen: &mut HashSet<String>,
        report: &mut dyn FnMut(usize),
    ) {
        if !formula.contains('!') && self.names.is_empty() && self.tables.is_empty() {
            return;
        }
        let (Ok(locale), Ok(language)) = (get_locale("en"), get_language("en")) else {
            return;
        };
        let mut pending = vec![formula];
        while let Some(formula) = pending.pop() {
            let mut lexer = Lexer::new(formula, LexerMode::A1, locale, language);
            let mut after_bracket = false;
            loop {
                let position = lexer.get_position();
                let token = lexer.next_token();
                if lexer.get_position() == position || matches!(token, TokenType::EOF) {
                    break;
                }
                let external =
                    std::mem::replace(&mut after_bracket, token == TokenType::RightBracket);
                let mut targets = None;
                match token {
                    TokenType::Reference {
                        sheet: Some(sheet), ..
                    }
                    | TokenType::Range {
                        sheet: Some(sheet), ..
                    } if !external => {
                        let sheet = sheet.to_lowercase();
                        if let Some(&index) = self.sheets.get(&sheet) {
                            targets = Some(index..=index);
                        } else if let Some((first, last)) = sheet.split_once(':')
                            && let (Some(&first), Some(&last)) =
                                (self.sheets.get(first), self.sheets.get(last))
                        {
                            targets = Some(first.min(last)..=first.max(last));
                        }
                    }
                    TokenType::StructuredReference { table_name, .. } => {
                        if let Some(&index) = self.tables.get(&table_name.to_lowercase()) {
                            targets = Some(index..=index);
                        }
                    }
                    TokenType::Ident(name)
                        if !self.names.is_empty()
                            && lexer.peek_token() != TokenType::LeftParenthesis =>
                    {
                        let name = name.to_lowercase();
                        if !names_seen.contains(&name)
                            && let Some(definition) = self
                                .names
                                .get(&(Some(source), name.clone()))
                                .or_else(|| self.names.get(&(None, name.clone())))
                        {
                            names_seen.insert(name);
                            pending.push(definition);
                        }
                    }
                    _ => {}
                }
                for index in targets.into_iter().flatten() {
                    if index != source && known.insert(index) {
                        report(index);
                    }
                }
            }
        }
    }
}
