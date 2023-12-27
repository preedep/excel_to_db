use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use csv::Writer;
use log::{debug, error, info};
use prettytable::{Cell, Row, Table};
use rusqlite::{Connection, named_params};
use rusqlite::types::Value;
use rustyline::DefaultEditor;
use rustyline::error::ReadlineError;
use serde::Deserialize;
use thousands::Separable;

#[derive(Debug, Deserialize)]
struct ExcelRow {
    service_name: String,
    #[serde(deserialize_with = "de_opt_f64")]
    average_response_time_95_ms: Option<f64>,
    count: Option<i64>,
    #[serde(deserialize_with = "de_opt_f64")]
    max_response_time_95_ms: Option<f64>,
    #[serde(deserialize_with = "de_opt_f64")]
    min_response_time_95_ms: Option<f64>,
}

#[derive(Parser)]
#[command(author, version, about, long_about = None)]
struct Cli {
    /// The file name of the excel
    #[arg(short = 'f')]
    file_name: String,
    /// The sheet name of the excel
    #[arg(short = 's')]
    sheet_name: String,
}

// Convert value cell to Some(f64) if float or int, else None
fn de_opt_f64<'de, D>(deserializer: D) -> Result<Option<f64>, D::Error>
    where
        D: serde::Deserializer<'de>,
{
    let data_type = calamine::DataType::deserialize(deserializer)?;
    if let Some(float) = data_type.as_f64() {
        Ok(Some(float))
    } else {
        Ok(None)
    }
}

fn load_excel(
    file_name: &String,
    sheet_name: &String,
) -> Result<Vec<ExcelRow>, calamine::XlsxError> {
    debug!(
        "Loading excel with file name: {} and sheet name: {}",
        file_name, sheet_name
    );

    use std::time::Instant;
    let now = Instant::now();

    let wb: Result<Xlsx<_>, calamine::XlsxError> = open_workbook(file_name);
    let wb = wb.map(|mut wb| {
        let res = wb.worksheet_range(sheet_name).map(|range| {
            //skip the first row , for header
            range
                .rows()
                .skip(1)
                .map(|row| ExcelRow {
                    service_name: row.get(0).unwrap().get_string().unwrap().to_string(),
                    average_response_time_95_ms: Some(
                        row.get(1).unwrap().get_float().unwrap_or(0.0),
                    ),
                    count: Some(row.get(2).unwrap().as_i64().unwrap_or(0)),
                    max_response_time_95_ms: Some(row.get(3).unwrap().get_float().unwrap_or(0.0)),
                    min_response_time_95_ms: Some(row.get(4).unwrap().get_float().unwrap_or(0.0)),
                })
                .collect::<Vec<ExcelRow>>()
        });
        res.unwrap()
    });
    let elapsed = now.elapsed();
    info!("Load Excel Elapsed: {:.2?}", elapsed);

    wb
}

fn import_database(
    connection: &mut Connection,
    excel_rows: &Vec<ExcelRow>,
) -> Result<(), rusqlite::Error> {
    use std::time::Instant;
    let now = Instant::now();

    let mut stmt = connection.prepare(
        "
        INSERT INTO excel_rows (
            service_name,
            average_response_time_95_ms,
            count,
            max_response_time_95_ms,
            min_response_time_95_ms
        )
        VALUES (
            :service_name,
            :average_response_time_95_ms,
            :count,
            :max_response_time_95_ms,
            :min_response_time_95_ms
        )
        ",
    )?;
    for row in excel_rows {
        stmt.execute(named_params! {
            ":service_name": &row.service_name,
            ":average_response_time_95_ms": &row.average_response_time_95_ms,
            ":count": &row.count,
            ":max_response_time_95_ms": &row.max_response_time_95_ms,
            ":min_response_time_95_ms": &row.min_response_time_95_ms
        })
            .map(|ret| {
                debug!("Insert excel row: {:?}", ret);
            })?
    }
    let elapsed = now.elapsed();
    info!("Import data Elapsed: {:.2?}", elapsed);

    Ok(())
}

fn main() {
    pretty_env_logger::init();
    let cli = Cli::parse();

    let mut connection = Connection::open_in_memory().unwrap();
    create_table(&cli, &mut connection);

    let mut rl = DefaultEditor::new().unwrap();
    #[cfg(feature = "with-file-history")]
    if rl.load_history("history.txt").is_err() {
        println!("No previous history.");
    }
    loop {
        let readline = rl.readline("[SQL] >> ");
        match readline {
            Ok(mut line) => {
                rl.add_history_entry(line.as_str())
                    .expect("add history entry error");
                let export_cli = line.clone();
                let export_cli = export_cli.split("|out=").last();
                if let Some(export) = export_cli {
                    info!("Require export: with parameter {}", export);
                    let mut line = line.replace("|out", "");
                    query_statement_and_display(&mut connection,
                                                &mut line, Some(export.to_string()));
                } else {
                    query_statement_and_display(&mut connection,
                                                &mut line,
                                                None);
                }
            }
            Err(ReadlineError::Interrupted) => {
                println!("CTRL-C");
                break;
            }
            Err(ReadlineError::Eof) => {
                println!("CTRL-D");
                break;
            }
            Err(err) => {
                println!("Error: {:?}", err);
                break;
            }
        }
    }
    #[cfg(feature = "with-file-history")]
    rl.save_history("history.txt");
}

fn create_table(cli: &Cli, mut connection: &mut Connection) {
    let res = connection.execute(
        "
        CREATE TABLE excel_rows (
            service_name TEXT NOT NULL,
            average_response_time_95_ms REAL NOT NULL,
            count INTEGER NOT NULL,
            max_response_time_95_ms REAL NOT NULL,
            min_response_time_95_ms REAL NOT NULL
        );
        CREATE UNIQUE INDEX idx_service_name
        ON excel_rows (service_name);
        ",
        (),
    );
    match res {
        Ok(_) => {
            info!("Create table excel_rows successfully");
            let excel_rows = load_excel(&cli.file_name, &cli.sheet_name);
            match excel_rows {
                Ok(excel_rows) => {
                    info!("Load excel successfully");
                    let res = import_database(&mut connection, &excel_rows);
                    match res {
                        Ok(_) => info!("Import excel rows successfully"),
                        Err(e) => error!("Import excel rows error: {}", e),
                    }
                }
                Err(e) => error!("Load excel error: {}", e),
            }
        }
        Err(e) => error!("Create Table Error: {}", e),
    }
}

fn query_statement_and_display(connection: &mut Connection,
                               line: &mut String,
                               exported_file_name: Option<String>) {
    use std::time::Instant;
    let now = Instant::now();

    let mut statement = connection.prepare(&line);
    match statement {
        Ok(ref mut smt) => {
            let mut column_names: Vec<Cell> = Vec::new();
            //let mut datas : Vec<Vec<&str>> = Vec::new();
            //let mut row_columns: Vec<&str> = Vec::new();

            let columns = smt.column_names();
            for column in columns {
                column_names.push(Cell::new(column));
                //if exported_file_name.is_some() {
                //    row_columns.push(column.clone());
                //}
            }

            let mut rows = smt.query([]).unwrap();
            let mut table = Table::new();
            table.add_row(Row::new(column_names.clone()));
            while let Some(row) = rows.next().unwrap() {
                //println!("row: {:?}", row);
                let mut cells: Vec<Cell> = Vec::new();
                for i in 0..column_names.clone().len() {
                    let ret = row.get::<usize, Value>(i);
                    let _ret = match ret {
                        Ok(ret) => {
                            let x = match ret {
                                Value::Null => Cell::new("NULL"),
                                Value::Integer(i) => Cell::new(i.separate_with_commas().as_str()),
                                Value::Real(r) => Cell::new(r.separate_with_commas().as_str()),
                                Value::Text(t) => Cell::new(t.as_str()),
                                Value::Blob(_b) => Cell::new("BLOB"),
                            };
                            cells.push(x);
                        }
                        Err(e) => {
                            error!("Get value error: {}", e);
                            continue;
                        }
                    };
                }
                table.add_row(Row::new(cells));
            }
            table.printstd();

            // Create csv writer
            //let mut writer : Option<Writer<File>> = None;
            if let Some(file) = exported_file_name {
                let wtr = Writer::from_path(file.clone());
                match wtr {
                    Ok(mut wtr) => {
                        for i in 0..table.len() {
                            if let Some(row_item) = table.get_row(i) {
                                let mut row_columns: Vec<String> = Vec::new();
                                row_item.iter().for_each(|cell| {
                                    let cell_value = cell.get_content();
                                    row_columns.push(cell_value.replace(",", ""));
                                });
                                wtr.write_record(row_columns).expect("write record error");
                            }
                        }
                        info!("Export csv successfully at file {}", file.clone());
                    }
                    Err(e) => {
                        error!("Create csv writer error: {}", e);
                    }
                }
            }
        }
        Err(e) => {
            error!("Statement error: {}", e);
        }
    }
    let elapsed = now.elapsed();
    info!("Query and Display Elapsed: {:.2?}", elapsed);
}
