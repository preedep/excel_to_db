use log::debug;
use serde::Deserialize;

#[derive(Debug,Deserialize)]
struct ExcelRow {
    service_name: String,
    average_response_time_95_ms: f64,
    max_response_time_95_ms: f64,
    min_response_time_95_ms: f64,
    count: u64,
}

fn main() {
    pretty_env_logger::init();

    debug!("Hello, world!");
}
