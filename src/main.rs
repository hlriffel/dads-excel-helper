use std::collections::HashMap;
use std::fmt;
use std::fmt::Formatter;
use std::fs::{File};
use std::path::Path;
use std::str::{FromStr};

use calamine::{DataType, open_workbook, Reader, Xlsx};
use xlsxwriter::{Workbook};
use chrono::{Datelike, Duration, NaiveDate, NaiveDateTime};
use clap::Parser;
use regex::Regex;

const INVOICE_EMISSION_INDEX: i8 = 0;
const INVOICE_NUMBER_INDEX: i8 = 1;
const CLIENT_NAME_INDEX: i8 = 4;
const PAYMENT_INTERVAL_INDEX: i8 = 8;
const INVOICE_VALUE_INDEX: i8 = 10;

#[derive(Parser)]
struct Args {
    /// Caminho da planilha de origem
    #[clap(short, long)]
    input: String,

    /// Planilha de entrada
    #[clap(short, long, default_value = "VENDAS")]
    sheet: String,

    /// Caminho onde a planilha resultado sera salva
    #[clap(short, long)]
    output: String,
}

struct Invoice {
    emission_date: NaiveDate,
    number: f64,
    client: String,
    payment_interval: String,
    value: f64,
}

impl fmt::Display for Invoice {
    #[allow(unused)]
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        write!(f, "{},{},{},{},{}",
               &*self.emission_date.format("%m/%d/%Y").to_string(), &*self.number.to_string(),
               &*self.client, &*self.payment_interval, &*self.value.to_string())
    }
}

struct CommissionedInvoice {
    emission_date: NaiveDate,
    number: f64,
    client: String,
    installment_value: f64,
    commission_value: f64,
}

impl fmt::Display for CommissionedInvoice {
    #[allow(unused)]
    fn fmt(&self, f: &mut Formatter<'_>) -> fmt::Result {
        write!(f, "{},{},{},{},{}",
               &*self.emission_date.format("%m/%d/%Y").to_string(), &*self.number.to_string(),
               &*self.client, &*self.installment_value.to_string(), &*self.commission_value.to_string())
    }
}

fn main() {
    let args = Args::parse();
    let invoices: Vec<Invoice> = get_invoices(&args);
    let commissions_by_month = get_commissions_by_month(invoices);
    let mut ordered_months = commissions_by_month.keys().cloned()
        .collect::<Vec<String>>();

    ordered_months.sort_by(|a, b| {
        println!("{}, {}", a, b);

        let a_date = NaiveDate::parse_from_str(a, "%B %Y");
        let b_date = NaiveDate::parse_from_str(b, "%B %Y");

        a_date.cmp(&b_date)
    });

    create_commission_sheets(&args.output, ordered_months, &commissions_by_month);

    // for (month, commissions) in commissions_by_month {
    //     println!("\n\n{}", month);
    //
    //     for commission in commissions {
    //         println!("{}", commission);
    //     }
    // }
}

fn get_invoices(args: &Args) -> Vec<Invoice> {
    let mut invoices: Vec<Invoice> = Vec::new();
    let mut workbook: Xlsx<_> = open_workbook(&args.input).unwrap();

    if let Some(Ok(range)) = workbook.worksheet_range(&args.sheet) {
        for row in range.rows().skip(1) {
            let emission_date = match &row[INVOICE_EMISSION_INDEX as usize] {
                DataType::DateTime(f) => {
                    let unix_secs = (f - 25569.) * 86400.;
                    let secs = unix_secs.trunc() as i64;
                    let nsecs = (unix_secs.fract().abs() * 1e9) as u32;
                    let day_month_year_date = NaiveDateTime::from_timestamp(secs, nsecs).date();

                    Some(NaiveDate::from_ymd(day_month_year_date.year(),
                                             day_month_year_date.day(),
                                             day_month_year_date.month()))
                }
                DataType::String(s) => { Some(NaiveDate::parse_from_str(s, "%m/%d/%Y").unwrap()) }
                _ => None
            };

            if emission_date.is_some() {
                invoices.push(Invoice {
                    emission_date: emission_date.unwrap(),
                    number: row[INVOICE_NUMBER_INDEX as usize].get_float().unwrap_or(f64::from(0)),
                    client: String::from(row[CLIENT_NAME_INDEX as usize].get_string().unwrap_or("")),
                    payment_interval: String::from(row[PAYMENT_INTERVAL_INDEX as usize].get_string().unwrap_or("")),
                    value: row[INVOICE_VALUE_INDEX as usize].get_float().unwrap_or(f64::from(0)),
                });
            }
        }
    }

    invoices
}

fn get_commissions_by_month(invoices: Vec<Invoice>) -> HashMap<String, Vec<CommissionedInvoice>> {
    let mut invoices_by_month = HashMap::new();
    let installments_regex: Regex = Regex::new(r"^((\d{2,3}/?)+)").unwrap();
    let special_commission_regex: Regex = Regex::new(r"(\d)%").unwrap();

    for invoice in invoices {
        if invoice.payment_interval.trim() == "ANTECIPADO / A VISTA [2]"
            || !installments_regex.is_match(&invoice.payment_interval) {
            let next_month = (invoice.emission_date + Duration::days(30))
                .format("%B %Y").to_string();
            let commissioned_invoices = invoices_by_month
                .entry(next_month).or_insert(Vec::new());

            commissioned_invoices.push(CommissionedInvoice {
                emission_date: invoice.emission_date,
                number: invoice.number,
                client: invoice.client,
                installment_value: invoice.value,
                commission_value: (f64::from(7) * invoice.value) / f64::from(100),
            });
        } else {
            let intervals: Vec<&str> = installments_regex
                .captures(&invoice.payment_interval).unwrap().get(1)
                .map(|m| m.as_str().split("/")).unwrap().collect::<Vec<&str>>();
            let commission: f64 = if special_commission_regex.is_match(&invoice.payment_interval) {
                special_commission_regex.captures(&invoice.payment_interval).unwrap().get(1)
                    .map(|m| f64::from_str(m.as_str())).unwrap().unwrap()
            } else { 7. };

            for interval in &intervals {
                let days = i16::from_str(&interval).unwrap();
                let installment_month = (invoice.emission_date + Duration::days(days as i64))
                    .format("%B %Y").to_string();
                let installment_value = invoice.value / intervals.len() as f64;

                invoices_by_month.entry(installment_month).or_insert(Vec::new())
                    .push(CommissionedInvoice {
                        emission_date: invoice.emission_date,
                        number: invoice.number,
                        client: invoice.client.clone(),
                        installment_value,
                        commission_value: (commission * installment_value) / f64::from(100),
                    })
            }
        }
    }

    invoices_by_month
}

fn create_commission_sheets(output_sheet: &str, ordered_months: Vec<String>, commissions_by_month: &HashMap<String, Vec<CommissionedInvoice>>) {
    ensure_file_is_created(&output_sheet);

    let workbook = Workbook::new(output_sheet);

    for month in ordered_months {
        let commissions = commissions_by_month.get(&month).unwrap();
        let mut worksheet = workbook.add_worksheet(Some(&month)).unwrap();

        // headers
        worksheet.write_string(0, 0, "Emissão", None);
        worksheet.write_string(0, 1, "Nr. NF", None);
        worksheet.write_string(0, 2, "Cliente", None);
        worksheet.write_string(0, 3, "Vlr. Parcela", None);
        worksheet.write_string(0, 4, "Vlr. Comissão", None);

        // data
        for (index, commission) in commissions.iter().enumerate() {
            worksheet.write_string((index + 1) as u32, 0, &commission.emission_date.format("%m/%d/%Y").to_string(), None);
            worksheet.write_string((index + 1) as u32, 1, &commission.number.to_string(), None);
            worksheet.write_string((index + 1) as u32, 2, &commission.client, None);
            worksheet.write_string((index + 1) as u32, 3, &commission.installment_value.to_string(), None);
            worksheet.write_string((index + 1) as u32, 4, &commission.commission_value.to_string(), None);
        }
    }

    workbook.close();
}

fn ensure_file_is_created(output_sheet: &str) {
    let output_path = Path::new(output_sheet);

    if !output_path.exists() {
        match File::create(output_path) {
            Err(cause) => panic!("could not create {}: {}", output_sheet, cause),
            Ok(_) => println!("successfully created {}", output_sheet),
        }
    }
}
