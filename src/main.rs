use std::io;

use calamine::{open_workbook, Reader, Xlsx};
use clap::Parser;
use log::{debug, error, info};
use rusqlite::{Connection, named_params};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct ExcelRow {
    service_name: String,
    average_response_time_95_ms: f64,
    count: u64,
    max_response_time_95_ms: f64,
    min_response_time_95_ms: f64,
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
                    average_response_time_95_ms: row.get(1).unwrap().get_float().unwrap_or(0.0),
                    count: row.get(2).unwrap().as_i64().unwrap_or(0) as u64,
                    max_response_time_95_ms: row.get(3).unwrap().get_float().unwrap_or(0.0),
                    min_response_time_95_ms: row.get(4).unwrap().get_float().unwrap_or(0.0),
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
        }).map(|ret|{
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
    let res = connection.execute(
        "
        CREATE TABLE excel_rows (
            service_name TEXT NOT NULL,
            average_response_time_95_ms REAL NOT NULL,
            count INTEGER NOT NULL,
            max_response_time_95_ms REAL NOT NULL,
            min_response_time_95_ms REAL NOT NULL
        );
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

    let mut line = String::new();
    loop {
        println!("Please enter query statement > ");
        io::stdin().read_line(&mut line).unwrap();
        let mut stmt = connection.prepare(&line);
        match stmt {
            Ok(mut smt) => {
                let mut rows = smt.query([]).unwrap();
                while let Some(row) = rows.next().unwrap() {
                    println!("row: {:?}", row);
                }
            }
            Err(e) => {
                error!("Statement error: {}", e);
            }
        }
        line.clear();
    }
}
