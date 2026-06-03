# PowerShell helper to build the GUI into a single exe using PyInstaller
# Run in project folder where invoice_converter_gui.py and convert_xml_invoices_to_excel.py live.
# The spec already includes convert_xml_invoices_to_excel.py and mapping_template.yaml as bundled data.

pyinstaller --clean --noconfirm invoice_converter_gui.spec

Write-Host "Build complete. Check the dist\ folder for invoice_converter_gui.exe"
