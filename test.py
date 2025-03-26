import os
import pandas as pd
import logging
from pathlib import Path
from datetime import datetime


class FileRenamer:
    def __init__(self, shared_drive_path: str, excel_mapping_path: str, dry_run: bool = True, log_dir: str = './logs'):
        self.shared_drive_path = Path(shared_drive_path)
        self.excel_mapping_path = excel_mapping_path
        self.dry_run = dry_run
        self.file_mapping = {}
        self.rename_log = []
        self._setup_logging(log_dir)

    def _setup_logging(self, log_dir):
        os.makedirs(log_dir, exist_ok=True)
        log_file = Path(log_dir) / f"file_renamer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        logging.basicConfig(
            filename=log_file,
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        logging.info("Logger initialized.")

    def load_mapping(self):
        try:
            df = pd.read_excel(self.excel_mapping_path, engine='openpyxl')
            if 'CurrentFilename' not in df.columns or 'NewFilename' not in df.columns:
                raise ValueError("Excel must contain 'CurrentFilename' and 'NewFilename' columns.")

            self.file_mapping = dict(zip(df['CurrentFilename'], df['NewFilename']))
            logging.info(f"Loaded {len(self.file_mapping)} filename mappings.")
        except Exception as e:
            logging.error(f"Failed to load Excel file: {e}")
            raise

    def rename_files(self):
        if not self.shared_drive_path.exists():
            logging.error(f"Shared drive path not found: {self.shared_drive_path}")
            return

        logging.info(f"{'Dry-run' if self.dry_run else 'Live'} mode enabled.")
        renamed_count = 0

        for file_path in self.shared_drive_path.rglob('*'):
            if not file_path.is_file():
                continue

            current_name = file_path.name
            if current_name in self.file_mapping:
                new_name = self.file_mapping[current_name]
                new_path = file_path.with_name(new_name)

                entry = {
                    "Original Path": str(file_path),
                    "New Path": str(new_path),
                    "Status": "",
                    "Error": ""
                }

                try:
                    if self.dry_run:
                        entry["Status"] = "Dry-run: Rename simulated"
                        logging.info(f"[Dry-run] Would rename: {current_name} -> {new_name}")
                    else:
                        os.rename(file_path, new_path)
                        renamed_count += 1
                        entry["Status"] = "Renamed"
                        logging.info(f"Renamed: {current_name} -> {new_name}")
                except Exception as e:
                    entry["Status"] = "Failed"
                    entry["Error"] = str(e)
                    logging.error(f"Failed to rename {current_name} to {new_name}: {e}")

                self.rename_log.append(entry)
            else:
                logging.debug(f"Skipped (not in Excel): {file_path.name}")

        logging.info(f"Renaming process complete. Total files renamed: {renamed_count}")

    def export_log_to_excel(self, output_dir: str = "./logs"):
        if not self.rename_log:
            logging.info("No log entries to export.")
            return

        os.makedirs(output_dir, exist_ok=True)
        report_path = Path(output_dir) / f"rename_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df_log = pd.DataFrame(self.rename_log)
        df_log.to_excel(report_path, index=False, engine='xlsxwriter')
        logging.info(f"Rename report exported to: {report_path}")

    def run(self):
        self.load_mapping()
        self.rename_files()
        self.export_log_to_excel()


if __name__ == "__main__":
    # Example usage
    SHARED_DRIVE_PATH = str(Path.home() / "SharedDrive")  # Example: "~/SharedDrive"
    EXCEL_MAPPING_PATH = "/Users/yourname/Documents/filename_mapping.xlsx"
    
    # Set dry_run=False to actually perform renames
    renamer = FileRenamer(SHARED_DRIVE_PATH, EXCEL_MAPPING_PATH, dry_run=True)
    renamer.run()

