# Process per UD table, streaming output
for ud in UD_TABLES:
    if cancel_event.is_set():
        status_q.put(("status", "Cancelled"))
        log_write(log_path, "Cancelled by user.")
        status_q.put(("done", "Cancelled"))
        return

    status_q.put(("status", f"Preparing {ud}..."))
    header_written = False
    header_cols = None
    out_filename = f"{prefix}{ud}.csv"
    out_path = output_folder / out_filename

    # If out_path exists from previous run, remove it to start fresh
    try:
        if out_path.exists():
            out_path.unlink()
    except Exception as e:
        log_write(log_path, f"Warning: couldn't remove existing {out_path}: {e}")

    # Iterate subfolders
    for sub in sorted(input_folder.iterdir()):
        if cancel_event.is_set():
            status_q.put(("status", "Cancelled"))
            log_write(log_path, "Cancelled by user.")
            status_q.put(("done", "Cancelled"))
            return

        if not sub.is_dir():
            continue

        # Find all files matching UDxx in this subfolder
        matching_files = [
            f for f in sorted(sub.iterdir())
            if f.is_file()
            and f.suffix.lower() in (".xlsx", ".csv")
            and "MASTER" not in f.name.upper()
            and ud in f.name.upper()
        ]

        # If no files matched, log and move to next subfolder
        if not matching_files:
            log_write(log_path, f"No file containing {ud} found in folder: {sub.name}")
            continue

        # Pick preferred file: .csv first, otherwise first match
        preferred_file = next((f for f in matching_files if f.suffix.lower() == ".csv"), matching_files[0])

        # Update UI
        status_q.put(("status", f"Processing {ud}: {preferred_file.name}"))

        # Read file
        try:
            if preferred_file.suffix.lower() == ".xlsx":
                df = pd.read_excel(preferred_file, engine="openpyxl", sheet_name=0, header=0, dtype=object)
            else:
                df = pd.read_csv(preferred_file)
        except Exception as e:
            log_write(log_path, f"ERROR reading {preferred_file}: {e}")
            status_q.put(("log", f"ERROR reading {preferred_file}: {e}"))
            processed += 1
            status_q.put(("progress", processed, total_files))
            continue

        # Skip if no data rows
        if df.shape[0] < 1:
            log_write(log_path, f"Skipped (no data rows): {preferred_file}")
            status_q.put(("log", f"Skipped (no data rows): {preferred_file}"))
            processed += 1
            status_q.put(("progress", processed, total_files))
            continue

        # Normalize columns to header from first processed file
        if header_cols is None:
            header_cols = list(df.columns)

        # Align columns
        if list(df.columns) != header_cols:
            new_df = pd.DataFrame(columns=header_cols)
            for c in df.columns:
                if c in header_cols:
                    new_df[c] = df[c]
            for c in header_cols:
                if c not in new_df.columns:
                    new_df[c] = pd.NA
            new_df = new_df[header_cols]
            df_to_write = new_df
        else:
            df_to_write = df[header_cols]

        # Write header if not yet written; append rows otherwise
        try:
            if not header_written:
                df_to_write.to_csv(out_path, index=False, header=True, encoding="utf-8")
                header_written = True
            else:
                df_to_write.to_csv(out_path, index=False, header=False, mode="a", encoding="utf-8")
        except Exception as e:
            log_write(log_path, f"ERROR writing to {out_path}: {e}")
            status_q.put(("log", f"ERROR writing to {out_path}: {e}"))

        processed += 1
        status_q.put(("progress", processed, total_files))

    # After finishing this UD, if no output file produced at all, log it
    if not out_path.exists():
        log_write(log_path, f"No valid files found for {ud} (no output produced)")
