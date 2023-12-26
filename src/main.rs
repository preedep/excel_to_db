use calamine::{open_workbook, Reader, Xlsx};
use log::{debug, error, info};
use serde::Deserialize;

#[derive(Debug, Deserialize)]
struct ExcelRow {
    service_name: String,
    average_response_time_95_ms: f64,
    count: u64,
    max_response_time_95_ms: f64,
    min_response_time_95_ms: f64,
}

fn main() {
    pretty_env_logger::init();

    use std::time::Instant;
    let now = Instant::now();
    const WORK_SHEET: &'static str = "Sheet1";
    let wb: Result<Xlsx<_>, calamine::XlsxError> = open_workbook("demo.xlsx");
    let wb = wb.map(|mut wb|{
        let res = wb.worksheet_range(WORK_SHEET)
            .map(|range| {
                //skip the first row , for header
                range.rows().skip(1).map(|row| {
                    ExcelRow {
                        service_name: row.get(0).unwrap().get_string().unwrap().to_string(),
                        average_response_time_95_ms: row.get(1).unwrap().get_float().unwrap_or(0.0),
                        count: row.get(2).unwrap().as_i64().unwrap_or(0) as u64,
                        max_response_time_95_ms: row.get(3).unwrap().get_float().unwrap_or(0.0),
                        min_response_time_95_ms: row.get(4).unwrap().get_float().unwrap_or(0.0),
                    }
                }).collect::<Vec<ExcelRow>>()
            });
        res.unwrap()
    });
    let elapsed = now.elapsed();
    info!("Load Excel Elapsed: {:.2?}", elapsed);

}
