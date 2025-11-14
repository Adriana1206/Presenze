@app.route("/submit", methods=['POST'])
def submit():
    nome = request.form.get("nome")
    cognome = request.form.get("cognome")
    mese_str = request.form.get("mese", '').lower()
    ferie_str = request.form.get("ferie", "")
    malattia_str = request.form.get("malattia", "")
    permessi_str = request.form.get("permessi", "")

    # --- Data Parsing ---
    try:
        giorni_ferie = set(map(int, ferie_str.split(','))) if ferie_str else set()
        giorni_malattia = set(map(int, malattia_str.split(','))) if malattia_str else set()
        
        permessi = {}
        if permessi_str:
            for item in permessi_str.replace(' ', '').split(','):
                if not item:
                    continue
                day_str, time_range = item.split(':', 1)
                day = int(day_str)
                start_time_str, end_time_str = time_range.split('-')
                
                start_time = datetime.strptime(start_time_str, '%H:%M')
                end_time = datetime.strptime(end_time_str, '%H:%M')
                
                duration = (end_time - start_time).total_seconds() / 3600
                permessi[day] = {"range": time_range, "hours": duration}

    except ValueError:
        return "Formato giorni o permessi non valido. Usare i formati corretti (es. Ferie: 1,5. Permessi: 13:10:00-11:30)", 400

    current_year = date.today().year
    month_number = italian_months.get(mese_str)
    if not month_number:
        return "Mese non valido. Per favore inserisci un mese in italiano.", 400

    _, num_days = calendar.monthrange(current_year, month_number)
    it_holidays = holidays.Italy(years=current_year)

    # Giorni settimana IT
    giorni_settimana_it = [
        "Lunedì", "Martedì", "Mercoledì",
        "Giovedì", "Venerdì", "Sabato", "Domenica"
    ]

    # --- Workbook Setup ---
    wb = Workbook()
    ws = wb.active
    ws.title = mese_str.capitalize()

    # --- Styles ---
    bold_font = Font(bold=True)
    light_red_fill = PatternFill(start_color="FFCCCB", end_color="FFCCCB", fill_type="solid")
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    light_gold_fill = PatternFill(start_color="EEE8AA", end_color="EEE8AA", fill_type="solid")
    blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
    light_purple_fill = PatternFill(start_color="E6E6FA", end_color="E6E6FA", fill_type="solid")

    # --- Header ---
    ws.append(["Nome", nome, "Cognome", cognome])
    ws['B1'].fill = green_fill
    ws['D1'].fill = green_fill
    ws.append([])

    header_titles = ["Giorno", "Orario Mattina", "Orario Pomeriggio", "Permessi"]
    ws.append(header_titles)
    for cell in ws[ws.max_row]:
        cell.font = bold_font

    # --- Calendar Body ---
    for day_num in range(1, num_days + 1):
        current_date = date(current_year, month_number, day_num)

        weekday_it = giorni_settimana_it[current_date.weekday()]
        date_str = f"{weekday_it}, {current_date.strftime('%d-%m-%Y')}"

        row_data = [date_str, "9:00 - 13:00", "14:00 - 18:00", ""]

        if current_date.weekday() >= 5 or current_date in it_holidays:
            row_data = [date_str, "", "", ""]
            ws.append(row_data)
            for cell in ws[ws.max_row]:
                cell.fill = light_red_fill
        
        elif day_num in giorni_ferie:
            row_data = [date_str, "FERIE", "", ""]
            ws.append(row_data)
            for cell in ws[ws.max_row]:
                cell.fill = light_gold_fill
            ws.cell(row=ws.max_row, column=2).font = bold_font

        elif day_num in giorni_malattia:
            row_data = [date_str, "MALATTIA", "", ""]
            ws.append(row_data)
            for cell in ws[ws.max_row]:
                cell.fill = blue_fill
            ws.cell(row=ws.max_row, column=2).font = bold_font

        elif day_num in permessi:
            permesso_info = permessi[day_num]
            row_data[3] = permesso_info["range"]
            ws.append(row_data)
            for cell in ws[ws.max_row]:
                cell.fill = light_purple_fill
        else:
            ws.append(row_data)
