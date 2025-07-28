import streamlit as st
import pandas as pd
import altair as alt # Import Altair

st.set_page_config(layout="wide") # Set wide layout for better use of space

st.title("Cocoa Trading Sheet Automation")
st.write("Automate cocoa trading processes and calculations from the Excel sheet.")

# Define the file path and relevant sheet names
FILE_PATH = r"C:\Users\Klaudia Gonciarz\Downloads\Cocoa Trading Sheet.xlsx"
SHEET_NAME_BEANS = "Costing Beans"
SHEET_NAME_PRODUCTS = "Costing Products"
SHEET_NAME_FREIGHT = "Freight & Dressing"
SHEET_NAME_VALO = "Valo Ori & Dest"
SHEET_NAME_FX_FIX = "Market & FX Fix"
SHEET_NAME_FX_LIVE = "Market & FX Live"
# Add other sheet names as needed based on full Excel analysis

@st.cache_data # Cache the data loading for performance
def load_excel_data(file_path, sheet_names):
    """Loads specified sheets from an Excel file into a dictionary of Dataframes."""
    dataframes = {}
    for sheet_name in sheet_names:
        try:
            # Use header=0, assuming the first row contains headers
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

            # No longer converting first column to string here as we are using header=0

            dataframes[sheet_name] = df
            st.sidebar.success(f"Successfully loaded sheet: '{sheet_name}'")
        except FileNotFoundError:
            st.sidebar.error(f"Error: File not found at {file_path}")
            dataframes[sheet_name] = pd.DataFrame() # Return empty DataFrame on error
            break # Stop loading other sheets if file is not found
        except Exception as e:
            st.sidebar.error(f"Error loading sheet '{sheet_name}': {e}")
            dataframes[sheet_name] = pd.DataFrame() # Return empty DataFrame on error
    return dataframes

@st.cache_data # Cache data processing results
def process_costing_beans(df_beans):
    """Processes the 'Costing Beans' DataFrame."""
    df_processed = pd.DataFrame()
    if not df_beans.empty:
        try:
            # Use the actual column names loaded with header=0
            # Identify relevant columns based on expected names
            relevant_cols = ['FX', 'VALUE DATE', 'FX RATE', 'Unnamed: 4']
            cols_to_select = [col for col in relevant_cols if col in df_beans.columns]

            if cols_to_select:
                df_processed = df_beans[cols_to_select].copy()
                # Explicitly convert columns to handle mixed types and ensure correct dtypes
                if 'FX' in df_processed.columns:
                    df_processed['FX'] = df_processed['FX'].astype(str)
                if 'VALUE DATE' in df_processed.columns:
                    df_processed['VALUE DATE'] = pd.to_datetime(df_processed['VALUE DATE'], errors='coerce')
                if 'FX RATE' in df_processed.columns:
                    df_processed['FX RATE'] = pd.to_numeric(df_processed['FX RATE'], errors='coerce')
                if 'Unnamed: 4' in df_processed.columns:
                    df_processed['Unnamed: 4'] = pd.to_numeric(df_processed['Unnamed: 4'], errors='coerce')

                subset_cols = [col for col in ['VALUE DATE', 'FX', 'FX RATE', 'Unnamed: 4'] if col in df_processed.columns]
                # Only dropna if the columns actually exist in the dataframe
                if subset_cols and not df_processed[subset_cols].empty:
                     df_processed.dropna(subset=subset_cols, inplace=True)
                elif subset_cols: # If columns exist but are empty after selection
                     df_processed = pd.DataFrame(columns=df_processed.columns) # Return empty with correct columns
                else: # If no relevant columns were found at all
                     df_processed = pd.DataFrame() # Return empty DataFrame


            else:
                st.warning("Required columns for processing 'Costing Beans' not found.")
                df_processed = pd.DataFrame() # Return empty DataFrame if key columns are missing

        except Exception as e:
            st.error(f"Error processing 'Costing Beans' sheet: {e}")
            df_processed = pd.DataFrame() # Return empty DataFrame on error
    return df_processed

@st.cache_data # Cache data processing results
def process_fx_data(df_fx):
    """Processes FX related DataFrames (Market & FX Fix/Live)."""
    df_processed = pd.DataFrame()
    if not df_fx.empty:
        try:
            # Use the actual column names loaded with header=0
            # Identify relevant columns based on expected names
            relevant_cols = ['Quote Table', 'Delivery', 'Last'] # Assuming these map to FX, VALUE DATE, FX RATE after header=0
            cols_to_select = [col for col in relevant_cols if col in df_fx.columns]

            if cols_to_select:
                df_processed = df_fx[cols_to_select].copy()
                # Rename columns to consistent names
                rename_map = {'Quote Table': 'FX', 'Delivery': 'VALUE DATE', 'Last': 'FX RATE'}
                df_processed.rename(columns=rename_map, inplace=True)

                # Explicitly convert columns to handle mixed types and ensure correct dtypes
                if 'FX' in df_processed.columns:
                    df_processed['FX'] = df_processed['FX'].astype(str)
                if 'VALUE DATE' in df_processed.columns:
                    df_processed['VALUE DATE'] = pd.to_datetime(df_processed['VALUE DATE'], errors='coerce')
                if 'FX RATE' in df_processed.columns:
                    df_processed['FX RATE'] = pd.to_numeric(df_processed['FX RATE'], errors='coerce')

                subset_cols = [col for col in ['VALUE DATE', 'FX', 'FX RATE'] if col in df_processed.columns]
                # Only dropna if the columns actually exist in the dataframe
                if subset_cols and not df_processed[subset_cols].empty:
                    df_processed.dropna(subset=subset_cols, inplace=True)
                elif subset_cols: # If columns exist but are empty after selection
                    df_processed = pd.DataFrame(columns=df_processed.columns) # Return empty with correct columns
                else: # If no relevant columns were found at all
                    df_processed = pd.DataFrame() # Return empty DataFrame

            else:
                st.warning("Required columns for processing FX data not found (assuming 'Quote Table', 'Delivery', 'Last').")
                df_processed = pd.DataFrame() # Return empty DataFrame if key columns is missing
        except Exception as e:
            st.error(f"Error processing FX sheet: {e}")
            df_processed = pd.DataFrame() # Return empty DataFrame on error
    return df_processed


@st.cache_data # Cache data processing results
def process_freight_data(df_freight):
    """Processes the 'Freight & Dressing' DataFrame."""
    df_processed = pd.DataFrame()
    if not df_freight.empty:
        try:
            # Use the actual column names loaded with header=0
            # Assuming column names are 'Origin', 'Destination', 'FreightCost' or similar after loading
            # Need to inspect df_freight.columns after loading with header=0 to confirm
            # For now, let's try more generic Unnamed columns as column names might not be clean
            relevant_cols = [df_freight.columns[0], df_freight.columns[1], df_freight.columns[5]] if df_freight.shape[1] > 5 else [] # Assuming first 2 and 6th column based on index
            if len(relevant_cols) == 3:
                df_processed = df_freight[relevant_cols].copy()

                # Rename columns for clarity
                df_processed.columns = ['Origin', 'Destination', 'FreightCost']

                # Explicitly convert columns to handle mixed types and ensure correct dtypes
                if 'Origin' in df_processed.columns:
                    df_processed['Origin'] = df_processed['Origin'].astype(str)
                if 'Destination' in df_processed.columns:
                    df_processed['Destination'] = df_processed['Destination'].astype(str)
                if 'FreightCost' in df_processed.columns:
                    df_processed['FreightCost'] = pd.to_numeric(df_processed['FreightCost'], errors='coerce')

                subset_cols = [col for col in ['Origin', 'Destination', 'FreightCost'] if col in df_processed.columns]
                # Only dropna if the columns actually exist in the dataframe
                if subset_cols and not df_processed[subset_cols].empty:
                     df_processed.dropna(subset=subset_cols, inplace=True)
                elif subset_cols: # If columns exist but are empty after selection
                     df_processed = pd.DataFrame(columns=df_processed.columns) # Return empty with correct columns
                else: # If no relevant columns were found at all
                     df_processed = pd.DataFrame() # Return empty DataFrame

            else:
                st.warning("Required columns for processing 'Freight & Dressing' not found (assuming first, second, and sixth columns).")
                df_processed = pd.DataFrame()

        except Exception as e:
            st.error(f"Error processing 'Freight & Dressing' sheet: {e}")
            df_processed = pd.DataFrame()
    return df_processed

@st.cache_data # Cache data processing results
def process_costing_products_data(df_products):
    """Processes the 'Costing Products' DataFrame."""
    df_processed = pd.DataFrame()
    if not df_products.empty:
        try:
            # Process using actual headers loaded with header=0
            df_processed = df_products.copy()

            # Attempt more robust type conversion for all columns
            for col in df_processed.columns:
                # Try converting to numeric first, then datetime, then fallback to string
                # Use a more cautious approach for numeric conversion
                try:
                    df_processed[col] = pd.to_numeric(df_processed[col], errors='coerce')
                except:
                    pass # Ignore error if numeric conversion fails

                # If conversion to numeric resulted in all NaNs, try datetime
                if df_processed[col].isna().all():
                    try:
                        df_processed[col] = pd.to_datetime(df_processed[col], errors='coerce')
                    except:
                        pass # Ignore error if datetime conversion fails

                # If still all NaNs (or not numeric/datetime), ensure it's string for display compatibility
                if df_processed[col].dtype == 'object':
                     df_processed[col] = df_processed[col].astype(str)


            # Add more specific processing here based on actual content and calculations needed
            # e.g., renaming specific columns based on content inspection if necessary

        except Exception as e:
            st.error(f"Error processing 'Costing Products' sheet: {e}")
            df_processed = pd.DataFrame()
    return df_processed

@st.cache_data # Cache data processing results
def process_valo_data(df_valo):
    """Processes the 'Valo Ori & Dest' DataFrame."""
    df_processed = pd.DataFrame()
    if not df_valo.empty:
        try:
            # Use the actual column names loaded with header=0
            # Assuming relevant columns are 'Buying Diff', 'Costings', 'Break Even', 'Selling Diff', 'Margin'
            relevant_cols = ['Buying Diff', 'Costings', 'Break Even', 'Selling Diff', 'Margin']
            cols_to_select = [col for col in relevant_cols if col in df_valo.columns]

            if cols_to_select:
                 df_processed = df_valo[cols_to_select].copy()

                 # Explicitly convert relevant calculation columns to numeric
                 calc_cols = ['Buying Diff', 'Costings', 'Break Even', 'Selling Diff', 'Margin']
                 for col in calc_cols:
                     if col in df_processed.columns:
                         df_processed[col] = pd.to_numeric(df_processed[col], errors='coerce')

                 # Drop rows where key calculation results are NaN (e.g., Break Even or Margin)
                 subset_cols_dropna = [col for col in ['Break Even', 'Margin'] if col in df_processed.columns]
                 if subset_cols_dropna and not df_processed[subset_cols_dropna].empty:
                      df_processed.dropna(subset=subset_cols_dropna, inplace=True)
                 elif subset_cols_dropna: # If columns exist but are empty after selection
                      df_processed = pd.DataFrame(columns=df_processed.columns)
                 else: # If no relevant columns were found at all
                      df_processed = pd.DataFrame()


            else:
                st.warning("Required columns for processing 'Valo Ori & Dest' not found.")
                df_processed = pd.DataFrame()

        except Exception as e:
            st.error(f"Error processing 'Valo Ori & Dest' sheet: {e}")
            df_processed = pd.DataFrame()
    return df_processed


# Load all necessary sheets
sheet_names_to_load = [SHEET_NAME_BEANS, SHEET_NAME_PRODUCTS, SHEET_NAME_FREIGHT, SHEET_NAME_VALO, SHEET_NAME_FX_FIX, SHEET_NAME_FX_LIVE]
excel_data = load_excel_data(FILE_PATH, sheet_names_to_load)

# Get DataFrames from loaded data (handle potential missing sheets)
df_beans_raw = excel_data.get(SHEET_NAME_BEANS, pd.DataFrame())
df_freight_raw = excel_data.get(SHEET_NAME_FREIGHT, pd.DataFrame())
df_valo_raw = excel_data.get(SHEET_NAME_VALO, pd.DataFrame())
df_costing_products_raw = excel_data.get(SHEET_NAME_PRODUCTS, pd.DataFrame())
df_fx_fix_raw = excel_data.get(SHEET_NAME_FX_FIX, pd.DataFrame())
df_fx_live_raw = excel_data.get(SHEET_NAME_FX_LIVE, pd.DataFrame())

# Process the loaded DataFrames
df_processed_beans = process_costing_beans(df_beans_raw)
df_processed_freight = process_freight_data(df_freight_raw)
df_processed_fx_fix = process_fx_data(df_fx_fix_raw)
df_processed_fx_live = process_fx_data(df_fx_live_raw) # Process live data similarly for now
df_processed_costing_products = process_costing_products_data(df_costing_products_raw) # Process products data
df_processed_valo = process_valo_data(df_valo_raw) # Process valo data


# --- Calculation Functions ---
def calculate_freight_cost(freight_df, origin, destination, quantity_mt):
    """
    Calculates the total freight cost for a given origin, destination, and quantity.
    Assumes freight_df has columns 'Origin', 'Destination', and 'FreightCost'.
    Looks for a matching Origin and Destination and multiplies the corresponding
    'FreightCost' by the quantity.
    """
    if freight_df.empty:
        return None, "Freight data not available."

    # Filter on case-insensitive exact match for origin and destination
    # Ensure columns exist before filtering
    if 'Origin' not in freight_df.columns or 'Destination' not in freight_df.columns or 'FreightCost' not in freight_df.columns:
         return None, "Required columns for freight calculation not found."

    matching_rows = freight_df[
        (freight_df['Origin'].astype(str).str.contains(origin, case=False, na=False)) &
        (freight_df['Destination'].astype(str).str.contains(destination, case=False, na=False))
    ]

    if matching_rows.empty:
        return None, f"No freight rate found for {origin} to {destination}."

    # Ensure 'FreightCost' is numeric and drop NaNs
    freight_rates = pd.to_numeric(matching_rows['FreightCost'], errors='coerce').dropna()

    if freight_rates.empty:
        return None, "Freight rate found but is not numeric."

    # Use the first valid freight rate found (simplification - could average or handle differently)
    freight_rate_per_unit = freight_rates.iloc[0]

    total_freight = freight_rate_per_unit * quantity_mt

    return total_freight, f"Calculated using rate {freight_rate_per_unit:.2f} per MT."


def perform_currency_conversion(df_fx, selected_fx, value_to_convert):
    """
    Performs currency conversion using the latest FX rate for the selected pair.
    Assumes df_fx has columns 'FX', 'VALUE DATE', and 'FX RATE'.
    """
    if df_fx.empty:
        return None, None, None, "FX data not available."

    # Ensure required columns exist
    if 'FX' not in df_fx.columns or 'VALUE DATE' not in df_fx.columns or 'FX RATE' not in df_fx.columns:
        return None, None, None, "Required columns for FX conversion not found."

    filtered_fx = df_fx[df_fx['FX'] == selected_fx].copy()

    if filtered_fx.empty:
        return None, None, None, f"No data available for the selected FX pair '{selected_fx}'."

    # Ensure 'VALUE DATE' is datetime and 'FX RATE' is numeric before sorting/calculation
    filtered_fx['VALUE DATE'] = pd.to_datetime(filtered_fx['VALUE DATE'], errors='coerce')
    filtered_fx['FX RATE'] = pd.to_numeric(filtered_fx['FX RATE'], errors='coerce')
    filtered_fx.dropna(subset=['VALUE DATE', 'FX RATE'], inplace=True) # Drop rows missing key info

    if filtered_fx.empty:
         return None, None, None, f"No valid date or FX rate data for the selected FX pair '{selected_fx}'."


    latest_fx_row = filtered_fx.sort_values(by='VALUE DATE', ascending=False).iloc[0]
    conversion_fx_rate = latest_fx_row['FX RATE']
    conversion_date = latest_fx_row['VALUE DATE'].date()

    # Ensure value_to_convert is numeric
    try:
        value_to_convert_numeric = float(value_to_convert)
    except (ValueError, TypeError):
        return None, None, None, "Invalid value to convert."


    converted_value = value_to_convert_numeric * conversion_fx_rate

    return converted_value, conversion_fx_rate, conversion_date, "Conversion successful."


def calculate_valuation(valo_df, buying_diff, costing):
    """
    Calculates valuation metrics (Break Even, Margin) based on Valo data.
    Assumes valo_df has columns 'Buying Diff', 'Costings', 'Break Even', 'Selling Diff', 'Margin'.
    This is a simplified calculation based on the structure, actual logic might be more complex.
    """
    if valo_df.empty:
        return None, None, "Valuation data not available."

    # Find relevant row based on Buying Diff and Costings (simplified)
    # In a real scenario, the lookup would likely be more complex based on contract terms, etc.
    # Let's find a row where 'Buying Diff' is close and use its 'Selling Diff' for a sample calculation

    # Find the row with the closest 'Buying Diff' to the input (simplified lookup)
    if 'Buying Diff' not in valo_df.columns or 'Selling Diff' not in valo_df.columns:
         return None, None, "Required columns for valuation calculation not found ('Buying Diff', 'Selling Diff')."

    # Ensure 'Buying Diff' is numeric
    valo_df_numeric_buying_diff = valo_df.copy()
    valo_df_numeric_buying_diff['Buying Diff'] = pd.to_numeric(valo_df_numeric_buying_diff['Buying Diff'], errors='coerce')
    valo_df_numeric_buying_diff.dropna(subset=['Buying Diff'], inplace=True)

    if valo_df_numeric_buying_diff.empty:
        return None, None, "Valuation data empty after processing 'Buying Diff'."

    closest_buying_diff_row = valo_df_numeric_buying_diff.iloc[(valo_df_numeric_buying_diff['Buying Diff'] - buying_diff).abs().argsort()[0]]

    # Assuming Break Even = Buying Diff + Costings (simplified)
    # Assuming Margin = Selling Diff - Break Even (simplified)
    # We will use the 'Selling Diff' from the closest row found for a sample Margin calculation

    selling_diff_from_sheet = pd.to_numeric(closest_buying_diff_row.get('Selling Diff'), errors='coerce')

    if pd.isna(selling_diff_from_sheet):
         return None, None, "Selling Diff from sheet is not numeric."


    calculated_break_even = buying_diff + costing # Use input costing
    calculated_margin = selling_diff_from_sheet - calculated_break_even

    return calculated_break_even, calculated_margin, f"Calculated using Selling Diff ({selling_diff_from_sheet:.2f}) from sheet."


def calculate_costing_products(products_df, input_params):
    """
    Placeholder function to calculate costing for products.
    Replace with actual logic based on 'Costing Products' sheet.
    input_params would be a dictionary of user inputs related to product costing.
    """
    if products_df.empty:
        return None, "Costing Products data not available.", products_df.head() # Return empty head

    # --- Placeholder Logic ---
    # This is where you would implement the actual calculations from the 'Costing Products' sheet.
    # This might involve:
    # - Looking up product-specific costs based on product type, origin, etc.
    # - Applying processing costs, packaging costs, etc.
    # - Using FX rates to convert costs to a base currency.
    # - Summing up various cost components to get a total product cost.
    # - Calculating margins or profitability for products.

    # For now, just return a placeholder message and the processed data head
    calculated_cost = None # Replace with actual calculated cost
    message = "Costing Products calculation logic needs to be implemented based on the sheet's formulas."

    return calculated_cost, message, products_df.head() # Return processed data head for inspection


# --- Streamlit UI Layout ---

st.sidebar.header("Settings and Inputs")

# Using tabs for different sections
tab_titles = ["FX Rates & Conversion", "Freight Calculation", "Costing Beans Data", "Valuation", "Costing Products", "Other Sheets Info"]
tabs = st.tabs(tab_titles)

# --- Tab: FX Rates & Conversion ---
with tabs[0]:
    st.header("FX Rates & Conversion")

    st.subheader("Market & FX Fix Data")
    if not df_processed_fx_fix.empty:
        st.write("Data from 'Market & FX Fix' sheet (Head):")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_processed_fx_fix.head().to_string())

        # Add FX Rate Line Chart
        st.subheader("Historical FX Rates (Market & FX Fix)")
        # Ensure 'FX' column exists before getting unique values
        if 'FX' in df_processed_fx_fix.columns:
            fx_pairs_for_chart = df_processed_fx_fix['FX'].unique()
            selected_fx_for_chart = st.selectbox("Select FX Pair to Visualize", fx_pairs_for_chart, key="fx_chart_selectbox") # Added key

            df_chart_data = df_processed_fx_fix[df_processed_fx_fix['FX'] == selected_fx_for_chart].copy()

            if not df_chart_data.empty and 'VALUE DATE' in df_chart_data.columns and 'FX RATE' in df_chart_data.columns:
                # Ensure data types are correct for Altair
                df_chart_data['VALUE DATE'] = pd.to_datetime(df_chart_data['VALUE DATE'])
                df_chart_data['FX RATE'] = pd.to_numeric(df_chart_data['FX RATE'], errors='coerce')
                df_chart_data.dropna(subset=['VALUE DATE', 'FX RATE'], inplace=True)

                if not df_chart_data.empty:
                    chart = alt.Chart(df_chart_data).mark_line().encode(
                        x=alt.X('VALUE DATE', title='Date'),
                        y=alt.Y('FX RATE', title=f'{selected_fx_for_chart} Rate'),
                        tooltip=[alt.Tooltip('VALUE DATE', title='Date'), 'FX RATE'] # Updated tooltip
                    ).properties(
                        title=f'Historical {selected_fx_for_chart} FX Rates'
                    ).interactive() # Make the chart interactive

                    st.altair_chart(chart, use_container_width=True)
                else:
                    st.warning("No valid data for plotting the selected FX pair.")
            else:
                st.warning("Insufficient data or missing required columns to plot FX rates.")
        else:
             st.warning("FX column not found in processed FX data for chart.")


        st.subheader("Currency Conversion using Market & FX Fix")
        # Ensure 'FX' column exists before getting unique values
        if 'FX' in df_processed_fx_fix.columns:
            fx_pairs_fix = df_processed_fx_fix['FX'].unique()
            selected_fx_fix = st.selectbox("Select FX Pair for Conversion (Fixed Rates)", fx_pairs_fix, key="fx_conversion_selectbox") # Added key

            value_to_convert_fix = st.number_input(
                f"Enter a value in the base currency of {selected_fx_fix} to convert (Fixed Rate):",
                value=0.0,
                format="%.2f",
                key="fx_conversion_value" # Added key
            )

            converted_value_fix, conversion_fx_rate_fix, conversion_date_fix, message_fix = perform_currency_conversion(
                df_processed_fx_fix, selected_fx_fix, value_to_convert_fix
            )

            if converted_value_fix is not None:
                 st.write(f"Using FX Rate **{conversion_fx_rate_fix:.4f}** from **{conversion_date_fix.strftime('%Y-%m-%d')}**")
                 st.success(f"Converted value in the quote currency: **{converted_value_fix:.2f}**")
            else:
                 st.warning(message_fix)
        else:
             st.warning("FX column not found in processed FX data for conversion.")

    else:
        st.warning("Could not load or process 'Market & FX Fix' data.")

    st.subheader("Market & FX Live Data (Placeholder)")
    # In a real application, this would likely involve fetching live data
    if not df_fx_live_raw.empty: # Display raw data as placeholder
        st.write("Preview of 'Market & FX Live' sheet (Head):")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_fx_live_raw.head().to_string())
        st.info("Live data integration would go here.")
    else:
        st.warning("Could not load 'Market & FX Live' data.")


# --- Tab: Freight Calculation ---
with tabs[1]:
    st.header("Freight Cost Calculation")

    if not df_processed_freight.empty:
        st.write("Inputs for Freight Calculation:")
        # Get unique values for origin and destination from the processed data
        # Ensure columns exist before getting unique values
        potential_origins = df_processed_freight['Origin'].dropna().unique() if 'Origin' in df_processed_freight.columns else []
        potential_destinations = df_processed_freight['Destination'].dropna().unique() if 'Destination' in df_processed_freight.columns else []


        selected_origin = st.selectbox("Select Origin", potential_origins, key="freight_origin_selectbox") if len(potential_origins) > 0 else st.text_input("Enter Origin", key="freight_origin_input")
        selected_destination = st.selectbox("Select Destination", potential_destinations, key="freight_destination_selectbox") if len(potential_destinations) > 0 else st.text_input("Enter Destination", key="freight_destination_input")
        quantity_mt = st.number_input("Enter Quantity (MT)", value=100.0, format="%.2f", key="freight_quantity_input")

        # Call the calculation function
        freight_cost, message = calculate_freight_cost(df_processed_freight, selected_origin, selected_destination, quantity_mt)

        if freight_cost is not None:
            st.success(f"Total Freight Cost: **{freight_cost:.2f}**")
            st.info(message)
        else:
            st.warning(message)

        st.subheader("Processed Freight Data Preview (Head)")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_processed_freight.head().to_string())
    else:
        st.warning("Could not load or process 'Freight & Dressing' data.")


# --- Tab: Costing Beans Data ---
with tabs[2]:
    st.header("Costing Beans Data")
    if not df_processed_beans.empty:
        st.subheader("Filtered Data (Costing Beans)")
        # UI elements for filtering (moved from sidebar to main area)
        # Ensure 'FX' column exists before getting unique values
        if 'FX' in df_processed_beans.columns:
            fx_pairs = df_processed_beans['FX'].unique()
            selected_fx = st.selectbox("Select FX Pair for Filtering", fx_pairs, key="beans_fx_selectbox")

            # Handle case where df_processed_beans['VALUE DATE'] might be empty after filtering/processing
            min_date = df_processed_beans['VALUE DATE'].min().date() if 'VALUE DATE' in df_processed_beans.columns and not df_processed_beans['VALUE DATE'].empty else pd.to_datetime('today').date()
            max_date = df_processed_beans['VALUE DATE'].max().date() if 'VALUE DATE' in df_processed_beans.columns and not df_processed_beans['VALUE DATE'].empty else pd.to_datetime('today').date()


            start_date, end_date = st.date_input(
                "Select Date Range for Filtering",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="beans_date_filter" # Added key
            )
            start_datetime = pd.to_datetime(start_date)
            end_datetime = pd.to_datetime(end_date)

            filtered_df_beans = df_processed_beans[df_processed_beans['FX'] == selected_fx].copy()
            final_filtered_df_beans = filtered_df_beans[
                (filtered_df_beans['VALUE DATE'] >= start_datetime) &
                (filtered_df_beans['VALUE DATE'] <= end_datetime)
            ].copy()

            st.write(f"Showing data for: **{selected_fx}** from **{start_date.strftime('%Y-%m-%d')}** to **{end_date.strftime('%Y-%m-%d')}**")
            # Use to_string() as a fallback for display if st.dataframe fails
            st.text(final_filtered_df_beans.head().to_string())

            # Automated Currency Conversion using Costing Beans FX Rate
            st.subheader("Currency Conversion using Costing Beans FX Rate")
            value_to_convert = st.number_input(
                f"Enter a value in the base currency of {selected_fx} to convert:",
                value=0.0,
                format="%.2f",
                key="beans_conversion_value" # Added key to avoid potential conflicts
            )

            converted_value, conversion_fx_rate, conversion_date, message = perform_currency_conversion(
                 final_filtered_df_beans, selected_fx, value_to_convert # Use the filtered data
            )

            if converted_value is not None:
                st.write(f"Using FX Rate **{conversion_fx_rate:.4f}** from **{conversion_date.strftime('%Y-%m-%d')}**")
                st.success(f"Converted value in the quote currency: **{converted_value:.2f}**")
            else:
                st.warning(message)
        else:
             st.warning("FX column not found in processed Costing Beans data for filtering/conversion.")


    else:
        st.warning("Could not load or process 'Costing Beans' data.")

# --- Tab: Valuation ---
with tabs[3]:
    st.header("Valuation (Origin & Destination)")
    if not df_processed_valo.empty:
        st.write("Inputs for Valuation:")
        # Get unique values for relevant columns if they exist
        potential_buying_diffs = df_processed_valo['Buying Diff'].dropna().unique() if 'Buying Diff' in df_processed_valo.columns else []
        potential_costings = df_processed_valo['Costings'].dropna().unique() if 'Costings' in df_processed_valo.columns else []

        # Provide input fields for Buying Diff and Costings
        input_buying_diff = st.number_input("Enter Buying Difference:", value=0.0, format="%.2f", key="valo_buying_diff")
        input_costing = st.number_input("Enter Total Costing:", value=0.0, format="%.2f", key="valo_costing")

        # Trigger calculation
        if st.button("Calculate Valuation", key="valo_calculate_button"):
            calculated_be, calculated_margin, message = calculate_valuation(df_processed_valo, input_buying_diff, input_costing)

            if calculated_be is not None and calculated_margin is not None:
                st.subheader("Calculation Results:")
                st.write(f"Calculated Break Even: **{calculated_be:.2f}**")
                st.write(f"Calculated Margin: **{calculated_margin:.2f}**")
                st.info(message)
            else:
                 st.warning(message)

        st.subheader("Processed Valuation Data Preview (Head)")
         # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_processed_valo.head().to_string())

    else:
        st.warning("Could not load or process 'Valo Ori & Dest' data.")

# --- Tab: Costing Products ---
with tabs[4]: # Adjusted index for the new tab
    st.header("Costing Products")
    st.write("This section will contain the logic and UI for calculating product costs.")

    if not df_processed_costing_products.empty:
        st.write("Inputs for Product Costing:")
        # Add input fields relevant to product costing based on the Excel sheet
        # Example: Product type, origin, quantity, processing parameters

        st.subheader("Processed Costing Products Data Preview (Head)")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_processed_costing_products.head().to_string())


        st.subheader("Product Cost Calculation (Placeholder)")
        st.write("Calculation logic will be implemented here.")

        # Example of how you might call the calculation function
        # input_params = {"product_type": selected_product, "quantity": input_quantity, ...}
        # calculated_cost, message, data_preview = calculate_costing_products(df_processed_costing_products, input_params)

        # if calculated_cost is not None:
        #     st.success(f"Calculated Product Cost: **{calculated_cost:.2f}**")
        # else:
        #     st.info(message) # Display message even if calculation is not fully implemented

    else:
        st.warning("Could not load or process 'Costing Products' data.")


# --- Tab: Other Sheets Info ---
with tabs[5]: # Adjusted index for the new tab
    st.header("Information from Other Sheets")
    st.write("This section provides a preview of data from other sheets for reference.")
    st.warning("Note: Displaying raw data from the Excel file might be limited due to formatting issues.")

    if not df_costing_products_raw.empty:
        st.subheader(f"'{SHEET_NAME_PRODUCTS}' Sheet Info (Head)")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_costing_products_raw.head().to_string())
        st.write("Info:")
        st.text(df_costing_products_raw.info())
    else:
         st.warning(f"Could not load '{SHEET_NAME_PRODUCTS}' data.")

    if not df_valo_raw.empty:
        st.subheader(f"'{SHEET_NAME_VALO}' Sheet Info (Head)")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_valo_raw.head().to_string())
        st.write("Info:")
        st.text(df_valo_raw.info())
    else:
         st.warning(f"Could not load '{SHEET_NAME_VALO}' data.")

    # Add other sheets here as needed
    if not df_fx_fix_raw.empty:
        st.subheader(f"'{SHEET_NAME_FX_FIX}' Sheet Info (Head)")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_fx_fix_raw.head().to_string())
        st.write("Info:")
        st.text(df_fx_fix_raw.info())
    else:
         st.warning(f"Could not load '{SHEET_NAME_FX_FIX}' data.")

    if not df_fx_live_raw.empty:
        st.subheader(f"'{SHEET_NAME_FX_LIVE}' Sheet Info (Head)")
        # Use to_string() as a fallback for display if st.dataframe fails
        st.text(df_fx_live_raw.head().to_string())
        st.write("Info:")
        st.text(df_fx_live_raw.info())
    else:
         st.warning(f"Could not load '{SHEET_NAME_LIVE}' data.")

# --- General Error Handling (moved to the end) ---
# The specific error handling within the load_excel_data function is usually sufficient.
# Any unhandled exceptions during the app execution will be displayed by Streamlit automatically.