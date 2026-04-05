# -*- coding: utf-8 -*-
# ImportConsultantSheetList.pushbutton / script.py
#
# Reads a CSV exported from the consultant sheet list dashboard page and
# adds/updates sheets in the model. Sheet List schedules are driven by
# ViewSheet elements — the schedule is populated automatically as sheets exist.

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, ViewSchedule,
    Transaction, ElementId, BuiltInCategory
)
import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Clipboard
from System.Diagnostics import Process
from pyrevit import forms, script
import csv, io

doc    = __revit__.ActiveUIDocument.Document
output = script.get_output()


# -- Step 1: How to get the CSV ------------------------------------------------
SHEET_LIST_URL = 'https://consultant-sheet-list.vercel.app/sheet-list-updater.html'

choice = forms.CommandSwitchWindow.show(
    ['Generate from PDF', 'Browse for CSV'],
    message='How would you like to provide the sheet list?'
)
if not choice:
    script.exit()

if choice == 'Generate from PDF':
    ready = forms.alert(
        'A browser window will open with the sheet list tool.\n\n'
        'Come back here once you have your CSV ready.',
        title='Opening Sheet List Tool',
        ok=True, cancel=True
    )
    if not ready:
        script.exit()
    Process.Start(SHEET_LIST_URL)
    confirmed = forms.alert(
        'The sheet list tool is now open in your browser.\n\n'
        '1. Drop your PDF\n'
        '2. Click  Get List\n'
        '3. Click  Download & Copy CSV\n'
        '   (this copies the sheet list to your clipboard)\n'
        '4. Come back here and click OK',
        title='Waiting for sheet list',
        ok=True, cancel=True
    )
    if not confirmed:
        script.exit()
    clip = Clipboard.GetText()
    if not clip.startswith('NUMBER,'):
        forms.alert(
            'The clipboard does not contain a sheet list.\n'
            'Please run the tool again and use Browse for CSV to locate your file.',
            title='Nothing found in clipboard'
        )
        script.exit()
    csv_source = io.StringIO(clip)

else:
    csv_path = forms.pick_file(file_ext='csv', title='Select Consultant Sheet List CSV')
    if not csv_path:
        script.exit()
    csv_source = open(csv_path, 'r')


# -- Step 2: Read CSV ----------------------------------------------------------
# Expected columns: NUMBER, SHEET NAME, DISCIPLINE, ORDER-MAJOR, ORDER-MINOR
csv_rows = []
with csv_source as f:
    reader = csv.DictReader(f)
    for row in reader:
        num   = (row.get('NUMBER')      or '').strip()
        title = (row.get('SHEET NAME')  or '').strip()
        disc  = (row.get('DISCIPLINE')  or '').strip()
        major = (row.get('ORDER-MAJOR') or '').strip()
        minor = (row.get('ORDER-MINOR') or '').strip()
        if not num or not title:
            continue
        try:
            csv_rows.append({
                'number': num, 'title': title, 'discipline': disc,
                'major': int(major), 'minor': int(minor)
            })
        except ValueError:
            csv_rows.append({
                'number': num, 'title': title, 'discipline': disc,
                'major': 0, 'minor': 0
            })

if not csv_rows:
    forms.alert('No valid rows found in CSV.', title='Nothing Found')
    script.exit()


# -- Step 3: Read existing sheets from the model -------------------------------
all_sheets   = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
model_sheets = {s.SheetNumber: s for s in all_sheets}

csv_map  = {r['number']: r for r in csv_rows}
to_add   = [r for r in csv_rows if r['number'] not in model_sheets]
to_update = [r for r in csv_rows if r['number'] in model_sheets
             and model_sheets[r['number']].Name != r['title']]


# -- Step 4: Preview -----------------------------------------------------------
if not to_add and not to_update:
    forms.alert('All sheets are already up to date.', title='No Changes')
    script.exit()

preview_items = []
for r in to_add:
    preview_items.append('+ ADD     {}    {}'.format(r['number'], r['title']))
for r in to_update:
    preview_items.append('~ UPDATE  {}    {}  ->  {}'.format(
        r['number'], model_sheets[r['number']].Name, r['title']))

confirmed = forms.alert(
    '\n'.join(preview_items),
    title='Review changes',
    ok=True, cancel=True
)
if not confirmed:
    script.exit()


# -- Step 5: Apply -------------------------------------------------------------
def set_params(sheet, data):
    p_disc  = sheet.LookupParameter('DISCIPLINE')
    p_major = sheet.LookupParameter('ORDER-MAJOR')
    p_minor = sheet.LookupParameter('ORDER-MINOR')
    if p_disc:  p_disc.Set(data['discipline'])
    if p_major: p_major.Set(data['major'])
    if p_minor: p_minor.Set(data['minor'])


with Transaction(doc, 'Import Consultant Sheet List') as t:
    t.Start()
    for data in to_add:
        s = ViewSheet.Create(doc, ElementId.InvalidElementId)
        s.SheetNumber = data['number']
        s.Name        = data['title']
        set_params(s, data)
    for data in to_update:
        s = model_sheets[data['number']]
        s.Name = data['title']
        set_params(s, data)
    t.Commit()

if to_add:
    output.print_md('## Added  —  {} sheet(s)'.format(len(to_add)))
    output.print_table(
        [[r['number'], r['title']] for r in to_add],
        columns=['Number', 'Title']
    )
if to_update:
    output.print_md('## Updated  —  {} sheet(s)'.format(len(to_update)))
    output.print_table(
        [[r['number'], r['title']] for r in to_update],
        columns=['Number', 'Title']
    )

output.print_md('---\n**Done.** {} added / {} updated.'.format(len(to_add), len(to_update)))
