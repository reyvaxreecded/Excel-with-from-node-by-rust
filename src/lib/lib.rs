use calamine::{Reader, open_workbook_auto};
use regex::Regex;
use std::path::Path;
use umya_spreadsheet::*;

#[derive(Debug, Clone)]
pub struct UpdateCell {
    cell: u32,
    value: String,
}

impl UpdateCell {
    pub fn new(cell: u32, value: String) -> Self {
        UpdateCell { cell, value }
    }
}

const BASE_PATH: &str = concat!(env!("CARGO_MANIFEST_DIR"), "/src/assets/excel_files/");

pub fn read_excel_file(file_name: &str) -> Result<(), Box<dyn std::error::Error>> {
    let path = format!("{}{}", BASE_PATH, file_name);
    println!("Reading Excel file: {}", path);
    println!("Path exists? {}", Path::new(&path).exists());

    let mut workbook = open_workbook_auto(&path)?;
    println!("Workbook opened successfully.");
    if let Some(Ok(range)) = workbook.worksheet_range_at(0) {
        for row in range.rows() {
            for cell in row {
                print!("{} ", cell);
            }
            println!();
        }
    } else {
        println!("No data found in the specified sheet.");
    }

    Ok(())
}

pub fn upsert_row(
    file_name: &str,
    sheet_name: &str,
    row: Vec<UpdateCell>,
) -> Result<(), Box<dyn std::error::Error>> {
    let path = format!("{}{}", BASE_PATH, file_name);
    println!("Upserting row in Excel file: {}", path);

    if !Path::new(&path).exists() {
        return Err(format!("Fichier non trouvé : {}", path).into());
    }

    let mut book = reader::xlsx::read(&path)?;
    let sheet = book.get_sheet_by_name_mut(sheet_name).ok_or(
        "Feuille
{sheet_name} introuvable",
    )?;

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

    writer::xlsx::write(&book, &path)?;
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
        let result = read_excel_file("example.xlsx");
        assert!(result.is_ok(), "Failed to read the Excel file.");
    }
}
