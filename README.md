# ExcelReportGenerator

Python • Pandas • OpenPyXL
A Python application that automatically generates polished Excel reports from CSV data—complete with pivot tables, charts, and summary statistics.



Features :
  1.📊 CSV to Excel Conversion: Transform raw CSV data into structured Excel workbooks with validation.
  2.🔢 Pivot Tables: Automatically generate pivot tables for summary analysis.
  3.📈 Data Visualization: Embed bar charts (or optionally line or pie charts) to highlight key metrics.
  4.📝 Summary Statistics: Compute mean, median, min, max, standard deviation, and count for insight into data.
  5.🎨 Professional Styling: Apply clean formatting (bold headers, cell colors, aligned columns, adaptive widths).
  6.🖥️ User-Friendly GUI: Simple interface built with Tkinter, featuring file selection and report generation dialogs.




📁 How It Works :
  
  1.Load CSV Data:
      Imports with pandas.read_csv()
      Validates file structure and handles malformed input gracefully.
  
  2.Pivot Table Generation:
      Detects numeric/categorical columns automatically
      Builds summary pivots (e.g. totals by category or date)
  
  3.Chart Creation:
      Uses openpyxl.chart.BarChart, LineChart, or PieChart to visualize pivot data
      Positions charts adjacent to pivot tables for clarity.
  
  4.Statistical Summaries:
      Calculates mean, median, range, standard deviation, and record count
      Presents in a tabular summary at the top or side of the report.
  
  5.Excel Styling & Templates:
      Formats headers, applies conditional formatting, auto-adjusts widths, and freezes panes
      Supports templated layout if you wish to embed content into branded Excel files
  
  6.GUI Interaction:
      Load CSV Button: launch file picker
      Generate Report Button: save dialog and execution
      Provides feedback messages for errors or notifications



🎯 Suitable For
  Ideal for automating report workflows in business, finance, sales tracking, and research:
      Reducing repetitive data consolidation tasks
      Creating fast, polished Excel reports with analysis baked in
      Turning raw datasets into presentation-ready formats



✅ Benefits
    1.🕒 Time Freedom: Automates a process that usually requires manual Excel work
    2.🔍 Clarity: Pivot tables and charts improve interpretability of data
    3.⚙️ Customization: Easily styled or templated to match branding or formatting requirements
    4.🖱️ Accessibility: GUI ensures usability even for non-technical users



🧰 Tools
    1.Python: Core language for logic and automation
    2.Pandas: Handles data import, cleaning, pivots, and stats
    3.OpenPyXL: Manages Excel generation, charts, and formatting
    4.Tkinter: Provides a minimal, accessible GUI for end users


