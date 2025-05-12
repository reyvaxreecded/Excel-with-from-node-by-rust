use rust_calamine::{UpdateCell, upsert_row};

fn main() {
    let file_name = "payplan2025VN.xlsx";
    let row = vec![
        UpdateCell::new(1, "Antony".to_string()),
        UpdateCell::new(2, "Hall".to_string()),
        UpdateCell::new(3, "10".to_string()),
        UpdateCell::new(4, "12".to_string()),
        UpdateCell::new(5, "15".to_string()),
        UpdateCell::new(6, "17".to_string()),
        UpdateCell::new(7, "20".to_string()),
        UpdateCell::new(8, "22".to_string()),
        UpdateCell::new(9, "25".to_string()),
        UpdateCell::new(10, "Oui".to_string()),
        UpdateCell::new(11, "Non".to_string()),
        UpdateCell::new(12, "40".to_string()),
        UpdateCell::new(13, "35".to_string()),
        UpdateCell::new(14, "30".to_string()),
        UpdateCell::new(20, "1.1".to_string()),
        UpdateCell::new(21, "1.2".to_string()),
        UpdateCell::new(22, "0.95".to_string()),
        UpdateCell::new(24, "1".to_string()),
        UpdateCell::new(25, "0.9".to_string()),
        UpdateCell::new(27, "30".to_string()),
    ];
    match upsert_row(file_name, "Payplan RENAULT VN", row) {
        Ok(_) => println!("Row upserted successfully."),
        Err(e) => eprintln!("Error upserting row: {}", e),
    }
}
