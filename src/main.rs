#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::fs::{self};
use std::io::prelude::*;
use std::ops::Not;
use std::sync::{Arc, Mutex};
use std::{env, fs::OpenOptions};

use calamine::{DataType, Reader, ToCellDeserializer, Xlsx, open_workbook};
use eframe::egui::{self, Layout};
use native_dialog::DialogBuilder;
use object::{Object, ObjectSection};
use rfd::FileDialog;

#[used]
#[unsafe(link_section = "inptdir")]
static mut INPUT_DIR_BYTES: [u8; 260] = [0; 260];

#[used]
#[unsafe(link_section = "outfil")]
static mut OUTPUT_FILE_BYTES: [u8; 260] = [0; 260];

// todo: move to ui, we'll also then probably embed these into the binary
// at which point, we might want to use a single section for all properties
const DATA_START_ID: &str = "Hole Number";
const DATA_END_ID: &str = "Sub-Totals";
const REMARKS_START_ID: &str = "Remarks";

#[derive(Default)]
struct App {
	input_dir: String,
	output_file: String,
	shared_state: Arc<Mutex<AppState>>,
}

#[derive(Default, Clone)]
struct AppState {
	input_dir: String,
	output_file: String,
}

impl App {
	fn new(input_dir: String, output_file: String) -> (App, Arc<Mutex<AppState>>) {
		let shared_state = Arc::new(Mutex::new(AppState {
			input_dir: input_dir.clone(),
			output_file: output_file.clone(),
		}));

		let app = App {
			input_dir,
			output_file,
			shared_state: shared_state.clone(),
		};

		(app, shared_state)
	}

	fn update_input_dir(&mut self, new_dir: String) {
		self.input_dir = new_dir.clone();
		if let Ok(mut state) = self.shared_state.lock() {
			state.input_dir = new_dir;
		}
	}

	fn update_output_file(&mut self, new_file: String) {
		self.output_file = new_file.clone();
		if let Ok(mut state) = self.shared_state.lock() {
			state.output_file = new_file;
		}
	}
}

impl eframe::App for App {
	fn update(&mut self, ctx: &eframe::egui::Context, _: &mut eframe::Frame) {
		egui::CentralPanel::default().show(ctx, |ui| {
			ui.with_layout(Layout::top_down_justified(egui::Align::Center), |ui| {
				ui.add_space(10.0);
				ui.horizontal(|ui| {
					if ui.button("Input Folder").clicked() {
						if let Some(folder) = FileDialog::new().pick_folder() {
							self.update_input_dir(folder.display().to_string());
						}
					}

					if self.input_dir.is_empty().not() {
						ui.label(&self.input_dir.to_string());
					}
				});

				ui.add_space(10.0);
				ui.horizontal(|ui| {
					if ui.button("Output File").clicked() {
						if let Some(file) = FileDialog::new().set_file_name("output.csv").save_file() {
							self.update_output_file(file.display().to_string());
						}
					}

					if self.output_file.is_empty().not() {
						ui.label(&self.output_file.to_string());
					}
				});

				ui.add_space(15.0);
				ui.separator();
				ui.add_space(15.0);

				ui.label("Selected Worksheets");
				ui.add_space(20.0);
				if &self.input_dir.len() > &0 {
					let scroll_area = egui::ScrollArea::vertical().max_height(120.0);
					scroll_area.show(ui, |ui| {
						if let Ok(entries) = std::fs::read_dir(&*self.input_dir) {
							for entry in entries.flatten() {
								let path = entry.path();
								if path.extension().and_then(|s| s.to_str()) == Some("xlsx") {
									let file_name = path
										.file_name()
										.and_then(|s| s.to_str())
										.unwrap_or_default();
									ui.label(file_name);
								}
							}
						} else {
							ui.label("No spreadsheets found in input directory.");
						}
					});
				}

				ui.add_space(20.0);
				let generate_button = egui::Button::new("Export");
				if self.output_file.is_empty().not() && self.input_dir.is_empty().not() {
					if ui.button("Generate").clicked() {
						match generate_output(self.input_dir.to_string(), self.output_file.to_string()) {
							Ok(_) => {
								DialogBuilder::message()
									.set_level(native_dialog::MessageLevel::Info)
									.set_title("Success")
									.set_text(format!("Aggregate data saved to: {}", self.output_file).as_str())
									.alert()
									.show()
									.unwrap();
							}
							Err(e) => {
								DialogBuilder::message()
									.set_level(native_dialog::MessageLevel::Error)
									.set_title("Error")
									.set_text(&format!("Failed to generate output: {}", e))
									.alert()
									.show()
									.unwrap();
							}
						}
					}
				} else {
					ui.add_enabled(false, generate_button);
				}
			});
		});
	}
}

fn generate_output(
	input_dir: String,
	output_file: String,
) -> Result<(), Box<dyn std::error::Error>> {
	if std::fs::exists(&output_file)? {
		std::fs::remove_file(&output_file)?;
	}

	let output_file = OpenOptions::new()
		.create(true)
		.append(true)
		.open(&output_file)?;

	let mut headers_set = false;
	for entry in std::fs::read_dir(input_dir)? {
		let path = entry?.path();
		if path.extension().and_then(|s| s.to_str()) != Some("xlsx") {
			continue;
		}

		let worksheet_name = path
			.file_stem()
			.and_then(|s| s.to_str())
			.unwrap_or_default();
		let mut workbook: Xlsx<_> = open_workbook(&path)?;

		if let Ok(r) = workbook.worksheet_range(worksheet_name) {
			let mut table_header_row: i32 = -1;
			let mut table_end_row: i32 = -1;
			let mut remarks_start_row: i32 = -1;

			let report_date = r
				.get((1, 0))
				.and_then(|data| data.as_string())
				.unwrap_or_default();
			for (row_idx, row) in r.rows().enumerate() {
				// might want to search more than just the first cell
				let first_cell = row.first().unwrap_or(&calamine::Data::Empty);
				if first_cell.as_string() == Some(DATA_START_ID.to_string()) {
					let sub_header_row = row_idx as i32 + 1;
					table_header_row = sub_header_row;

					if headers_set {
						continue;
					}

					let sub_headers = r.rows().nth(sub_header_row as usize).unwrap_or(&[]);
					let mut prev_main_header = String::new();
					let mut headers: Vec<_> = row
						.iter()
						.enumerate()
						.map(|(i, c)| {
							let main_header = c.as_string().unwrap_or_default();
							let main_header = if main_header.is_empty() {
								prev_main_header.as_str()
							} else {
								prev_main_header = main_header.to_string();
								&main_header
							};

							let sub_header = sub_headers
								.get(i)
								.and_then(|sc| sc.as_string())
								.unwrap_or_default();
							if sub_header.is_empty().not() {
								return format_header(format!("{}_{}", main_header, sub_header));
							}

							format_header(main_header.to_string())
						})
						.collect();

					headers.push("date".to_string());
					writeln!(&output_file, "{}", headers.join(","))?;
					headers_set = true;
					continue;
				}

				if first_cell.as_string() == Some(DATA_END_ID.to_string()) {
					table_end_row = row_idx as i32;
					continue;
				}

				if row_idx as i32 <= table_header_row {
					continue;
				}

				if remarks_start_row < 0 && first_cell.as_string() == Some(REMARKS_START_ID.to_string()) {
					remarks_start_row = row_idx as i32;
					continue;
				}

				if table_header_row > 0 && table_end_row < 0 {
					if ToCellDeserializer::is_empty(first_cell) {
						continue;
					}

					let row_data: Vec<_> = row.iter().map(|c| c.to_string()).collect();
					write!(&output_file, "{}", row_data.join(","))?;
					writeln!(&output_file, ",{}", report_date)?;
				}
			}
		}
	}

	Ok(())
}

fn format_header(header: String) -> String {
	header
		.trim()
		.to_lowercase()
		.replace(" ", "_")
		.replace("-", "_")
		.replace("\n", "_")
}

fn update_binary(sections: &[(&str, &str)]) {
	let exe_path = env::current_exe().unwrap();
	let tmp = exe_path.with_extension("tmp");
	fs::copy(&exe_path, &tmp).unwrap();

	let file = OpenOptions::new()
		.read(true)
		.write(true)
		.open(&tmp)
		.unwrap();
	let mut buf = unsafe { memmap2::MmapOptions::new().map_mut(&file) }.unwrap();
	let mut section_updates = Vec::new();
	{
		let parsed_file = object::File::parse(&*buf).unwrap();
		for (section_name, data) in sections {
			if let Some((offset, size)) = get_section(&parsed_file, section_name) {
				let section_data = &buf[offset as usize..(offset + size) as usize];
				if data.as_bytes() == section_data {
					continue;
				}

				section_updates.push((offset as usize, size as usize, *data));
			}
		}
	}

	if section_updates.is_empty().not() {
		for (offset, size, data) in &section_updates {
			buf[*offset..(*offset + *size)].fill(0);
			buf[*offset..*offset + data.len()].copy_from_slice(data.as_bytes());
		}

		let perms = fs::metadata(&exe_path).unwrap().permissions();
		let old = env::temp_dir().join(exe_path.with_extension("old"));

		fs::set_permissions(&tmp, perms.clone()).unwrap();
		// can't just overwrite running exe on windows, so move/rename
		// to temp and then rename back
		fs::rename(&exe_path, &old).unwrap();
		fs::rename(&tmp, &exe_path).unwrap();
		fs::set_permissions(&exe_path, perms).unwrap();
	} else {
		fs::remove_file(&tmp).unwrap();
	}
}

fn get_section(file: &object::File, name: &str) -> Option<(u64, u64)> {
	for section in file.sections() {
		match section.name() {
			Ok(n) if n == name => {
				return section.file_range();
			}
			_ => {}
		}
	}
	None
}

fn main() -> eframe::Result {
	// clean up old temp file
	let old_path = env::current_exe().unwrap().with_extension("old");
	let old_path = env::temp_dir().join(old_path);
	if old_path.exists() {
		let _ = fs::remove_file(old_path);
	}

	// read configured paths from binary sections
	// blame: https://blog.dend.ro/self-modifying-rust/
	let input_dir_bytes = unsafe { INPUT_DIR_BYTES };
	let output_file_bytes = unsafe { OUTPUT_FILE_BYTES };

	let input_dir_string = String::from_utf8_lossy(&input_dir_bytes);
	let input_dir = input_dir_string.trim_end_matches(char::from(0)).to_owned();
	let output_file_string = String::from_utf8_lossy(&output_file_bytes);
	let output_file = output_file_string
		.trim_end_matches(char::from(0))
		.to_owned();

	let options = eframe::NativeOptions {
		viewport: egui::ViewportBuilder::default()
			.with_inner_size([320.0, 480.0])
			.with_min_inner_size([320.0, 480.0]),
		..Default::default()
	};

	let (app, shared_state) = App::new(input_dir, output_file);
	let native_result = eframe::run_native("oxide", options, Box::new(|_cc| Ok(Box::new(app))));
	let final_state = shared_state.lock().unwrap().clone();
	let sections = [
		("inptdir", &final_state.input_dir as &str),
		("outfil", &final_state.output_file as &str),
	];

	update_binary(&sections);
	native_result
}
