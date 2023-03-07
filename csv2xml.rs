// Importing necessary modules
use std::error::Error;
use std::fs::File;
use std::io::prelude::*;
use std::path::Path;

// Main function of the program
fn main() -> Result<(), Box<dyn Error>> {
    // Path of the CSV file
    let csv_path = Path::new("input.csv");

    // Check if the CSV file exists
    if !csv_path.exists() {
        return Err("The CSV file does not exist.".into());
    }

    // Read the contents of the CSV file
    let mut csv_content = String::new();
    let mut csv_file = File::open(csv_path)?;
    csv_file.read_to_string(&mut csv_content)?;

    // Parse the contents of the CSV file
    let mut rows: Vec<Vec<String>> = vec![];
    for line in csv_content.lines() {
        rows.push(line.split(",").map(|s| s.to_string()).collect());
    }

    // Create a new Excel file
    let excel_path = Path::new("output.xml");
    let mut excel_file = File::create(excel_path)?;

    // Write the contents of the Excel file
    writeln!(excel_file, "<?xml version=\"1.0\"?>")?;
    writeln!(excel_file, "<?mso-application progid=\"Excel.Sheet\"?>")?;
    writeln!(excel_file, "<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"")?;
    writeln!(excel_file, " xmlns:o=\"urn:schemas-microsoft-com:office:office\"")?;
    writeln!(excel_file, " xmlns:x=\"urn:schemas-microsoft-com:office:excel\"")?;
    writeln!(excel_file, " xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"")?;
    writeln!(excel_file, " xmlns:html=\"http://www.w3.org/TR/REC-html40/\">")?;
    writeln!(excel_file, "<Worksheet ss:Name=\"Sheet1\">")?;
    writeln!(excel_file, "<Table>")?;
    for row in &rows {
        writeln!(excel_file, "<Row>")?;
        for col in row {
            writeln!(excel_file, "<Cell><Data ss:Type=\"String\">{}</Data></Cell>", col)?;
        }
        writeln!(excel_file, "</Row>")?;
    }
    writeln!(excel_file, "</Table>")?;
    writeln!(excel_file, "</Worksheet>")?;
    writeln!(excel_file, "</Workbook>")?;

    println!("The XML file has been successfully created.");

    std::io::stdin().read_line(&mut String::new()).unwrap();

    Ok(())
}
