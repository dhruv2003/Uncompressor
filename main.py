import os
import zipfile
import rarfile
from pathlib import Path

# === CONFIGURATION ===
zip_folder = r"C:\190 Crore Database"
output_excel_folder = r"C:\master sheets"

os.makedirs(output_excel_folder, exist_ok=True)

# === LOGGING ===
log_path = os.path.join(output_excel_folder, "extract_log.txt")
log = open(log_path, "w", encoding="utf-8")

def log_print(msg):
    print(msg)
    log.write(msg + "\n")

# === SUMMARY COUNTERS ===
total_archives = 0
archives_with_excels = 0
total_excels_found = 0

# === START PROCESSING ===
log_print("🔍 Starting archive scan...\n")

for root, dirs, files in os.walk(zip_folder):
    for file in files:
        archive_path = os.path.join(root, file)
        ext = file.lower().strip().split('.')[-1]

        if ext not in ('zip', 'rar'):
            continue  # skip non-archives

        total_archives += 1
        log_print(f"📁 Processing: {archive_path}")

        try:
            if ext == 'zip':
                archive = zipfile.ZipFile(archive_path, 'r')
                members = archive.namelist()
                opener = archive.open
            elif ext == 'rar':
                archive = rarfile.RarFile(archive_path, 'r')
                members = archive.namelist()
                opener = archive.open

            found_excel = False

            for member in members:
                if member.lower().endswith(('.xls', '.xlsx')):
                    found_excel = True
                    total_excels_found += 1

                    with opener(member) as source_file:
                        original_name = os.path.basename(member)
                        if not original_name:
                            continue

                        save_path = os.path.join(output_excel_folder, original_name)

                        counter = 1
                        while os.path.exists(save_path):
                            name, extn = os.path.splitext(original_name)
                            new_name = f"{name}_{counter}{extn}"
                            save_path = os.path.join(output_excel_folder, new_name)
                            counter += 1

                        with open(save_path, 'wb') as out_file:
                            out_file.write(source_file.read())

                        log_print(f"   ✅ Extracted: {original_name}")

            if not found_excel:
                log_print("   ⚠️ No Excel files found.")
            else:
                archives_with_excels += 1

            archive.close()

        except (zipfile.BadZipFile, rarfile.Error) as e:
            log_print(f"   ❌ Error reading archive: {e}")

log_print("\n✅ Extraction complete.")
log_print(f"📦 Total archives scanned: {total_archives}")
log_print(f"📄 Archives with Excel files: {archives_with_excels}")
log_print(f"📊 Total Excel files extracted: {total_excels_found}")
log.close()
