#![deny(clippy::all)]

use calamine::{Reader, open_workbook_auto};
use napi_derive::napi;
use regex::Regex;
use std::path::Path;
use umya_spreadsheet::*;

#[napi(object)]
#[derive(Debug, Clone)]
pub struct UpdateCell {
    pub cell: u32,
    pub value: String,
}

impl UpdateCell {
    pub fn new(cell: u32, value: String) -> Self {
        UpdateCell { cell, value }
    }
}



#[napi(js_name = "read_excel_file")]
pub fn read_excel_file(path_to_file: String, sheet_name: String) -> Result<Vec<Vec<String>>, napi::Error> {
    println!("Reading Excel file: {}", path_to_file);
    println!("Path exists? {}", Path::new(&path_to_file).exists());

    let mut workbook = match open_workbook_auto(&path_to_file) {
        Ok(workbook) => workbook,
        Err(e) => {
            println!("Error opening workbook: {}", e);
            return Err(napi::Error::new(
                napi::Status::GenericFailure,
                format!("Error opening workbook: {}", e),
            ));
        }
    };
    println!("Workbook opened successfully.");
    let sheet_names = workbook.sheet_names().to_vec();
    println!("Sheet names: {:?}", sheet_names);
    println!("Reading sheet: {}", sheet_name);
    let range = workbook.worksheet_range(&sheet_name).map_err(|e| {
        napi::Error::new(
            napi::Status::GenericFailure,
            format!("Error reading sheet: {}", e),
        )
    })?;

    Ok(range
        .rows()
        .map(|row| {
            row.iter()
                .map(|cell| cell.to_string())
                .collect::<Vec<String>>()
        })
        .collect())
}

#[napi(js_name = "upsert_row")]
pub fn upsert_row(
    path_to_file: String,
    sheet_name: String,
    row: Vec<UpdateCell>,
) -> Result<(), napi::Error> {
    println!("Upserting row in Excel file: {}", &path_to_file);

    if !Path::new(&path_to_file).exists() {
        return Err(napi::Error::new(
            napi::Status::GenericFailure,
            format!("File not found: {}", &path_to_file),
        ));
    }

    let mut book = reader::xlsx::read(&path_to_file).map_err(|e| {
        napi::Error::new(
            napi::Status::GenericFailure,
            format!("Error reading sheet: {}", e),
        )
    })?;
    let sheet = book.get_sheet_by_name_mut(&sheet_name).ok_or_else(|| {
        napi::Error::new(
            napi::Status::GenericFailure,
            format!("Sheet not found: {}", sheet_name),
        )
    })?;

    let mut found = false;

    for row_idx in 1..=sheet.get_highest_row() {
        let cell_value = sheet.get_formatted_value((1, row_idx));

        if cell_value == row[0].value {
            println!("Ligne trouvée à l'index {}. Mise à jour...", row_idx);
            for value in row.iter() {
                sheet
                    .get_cell_mut((value.cell, row_idx))
                    .set_value(value.value.clone());
            }
            found = true;
            break;
        }
    }

    if !found {
        let new_row_idx = sheet.get_highest_row() + 1;
        let previous_row_idx = new_row_idx - 1;
        let sheet_clone = &mut sheet.clone();
        println!(
            "Aucune correspondance trouvée. Ajout d'une nouvelle ligne à l'index {}.",
            new_row_idx
        );
        sheet.insert_new_row(&new_row_idx, &1);
        for value in row.iter() {
            let style = sheet_clone
                .get_cell_mut((value.cell, previous_row_idx))
                .get_style();
            sheet
                .get_cell_mut((value.cell, new_row_idx))
                .set_style(style.clone())
                .set_value(value.value.clone());
        }

        let specified_columns: Vec<u32> = row.iter().map(|c| c.cell).collect();
        let highest_column = sheet.get_highest_column();
        let sheet_clone = sheet.clone();

        let formulas_to_copy: Vec<(u32, &str)> = (1..=highest_column)
            .filter(|col| !specified_columns.contains(col))
            .filter_map(|col| {
                let cell = sheet_clone.get_cell((col, previous_row_idx))?;
                if cell.is_formula() {
                    let formula = cell.get_formula();
                    println!("Formule trouvée dans la colonne {}: {}", col, formula);
                    Some((col, formula))
                } else {
                    None
                }
            })
            .collect();

        for (col, formula) in formulas_to_copy {
            println!(
                "Copie de la formule {} de la colonne {} vers la nouvelle ligne {}",
                formula, col, new_row_idx
            );
            let sheet_clone = &mut sheet.clone();
            let style = sheet_clone
                .get_cell_mut((col, previous_row_idx))
                .get_style();
            sheet
                .get_cell_mut((col, new_row_idx))
                .set_style(style.clone())
                .set_formula(adjust_formula_references(
                    formula,
                    previous_row_idx,
                    new_row_idx,
                ));
        }
    }

    writer::xlsx::write(&book, &path_to_file).map_err(|e| {
        napi::Error::new(
            napi::Status::GenericFailure,
            format!("Error writing to file: {}", e),
        )
    })?;
    println!("Fichier sauvegardé avec succès.");

    Ok(())
}

fn adjust_formula_references(formula: &str, old_row: u32, new_row: u32) -> String {
    let pattern = format!(r"\b([A-Z]{{1,3}})({})\b", old_row);
    let re = Regex::new(&pattern).unwrap();
    re.replace_all(formula, |caps: &regex::Captures| {
        format!("{}{}", &caps[1], new_row)
    })
    .to_string()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    #[should_panic]
    fn test_read_excel_file() {
        let result = read_excel_file("example.xlsx".to_string());
        assert!(result.is_ok(), "Failed to read the Excel file.");
    }
}
