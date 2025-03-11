# SheetSherpa 
A powerful web-based tool for exploring and analyzing Excel/CSV files with an intuitive Streamlit interface. Upload your file, apply filters, and visualize data instantly.

### Key Features:
- **Interactive Filtering**:  
  - Numeric ranges, date ranges, and multi-select filters for columns.  
  - Advanced search with regex, AND/OR logic, and keyword operators (`&`, `|`, `!`).  
  - Real-time filtering with caching for performance.  

- **Data Manipulation**:  
  - Clean data by removing duplicates/missing values.  
  - Rename columns and apply custom conditions.  
  - Create pivot tables and aggregate data (sum, mean, count, etc.).  

- **Visualization**:  
  - Generate charts: bar, line, scatter, box plots, histograms, heatmaps, and time series.  
  - Interactive plots with Altair for seamless exploration.  

- **Exports**:  
  Download filtered data as Excel, CSV, or JSON files.  

- **Summary Stats**:  
  Automatically compute statistics for numeric/datetime columns.  

### Installation and Running the app:

#### Install Dependencies
Run the following command in your terminal to install required packages:

```bash
pip install -r requirements.txt
```

#### Run the App

Execute the `streamlit` app:

```bash
streamlit run sheet_sherpa.py
```

This will then open the app in your browser at [`streamlit`'s default local server](http://localhost:8501)

### How to Use the app:
1. Upload an Excel (`.xlsx`) or CSV file via the sidebar.  
2. Choose the sheet (if multi-sheet) and header row.  
3. Apply filters, search terms, or aggregation settings.  
4. Explore results, charts, and statistics in real-time.  

Built with 🚀Streamlit, Pandas, and Altair for fast, interactive data analysis.  

---

**Try it now!**  
Upload your file and start analyzing!
