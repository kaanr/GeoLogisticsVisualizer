import pandas as pd
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
import folium
from folium.plugins import MarkerCluster
from bs4 import BeautifulSoup
import re
import os
from datetime import datetime
import xlsxwriter

import pandas as pd
from bs4 import BeautifulSoup
import re

class DataProcessor:
    def __init__(self, input_filename):
        self.filename = input_filename
        self.df = None # Main DataFrame for operations like mapping
        self.all_areas_df = None # DataFrame for aggregation and Excel export

    def read_and_parse_html(self):
        # Read HTML file and parse it
        with open(self.filename, 'r', encoding='utf-8') as file:
            html_content = file.read()
        soup = BeautifulSoup(html_content, 'lxml')
        # Find the table by its class
        table = soup.find('table', class_='print-table-shipping-list')
        # Initialize a list to store each row's data
        data = []
        # Extract headings
        headings = [th.get_text(strip=True) for th in table.find('thead').find_all('th')]
        # Iterate through each row in the table (excluding the header)
        for row in table.find('tbody').find_all('tr'):
            # Extract each cell in the row
            cells = row.find_all(['th', 'td'])
            row_data = [cell.get_text(strip=True) for cell in cells]
            data.append(row_data)
        # Convert to DataFrame for easier manipulation and display
        self.df = pd.DataFrame(data, columns=headings)
        self.clean_and_transform_data()

    def clean_and_transform_data(self):
        # Renaming and cleaning specific columns
        self.df.rename(columns={'Адрес доставки\xa0/приема': 'Адрес доставки'}, inplace=True)

        # Using regex to clean and extract parts of addresses
        pattern_postal_code_region = r'^\d+,\s[^\,]+,\s+'
        self.df['Адрес доставки'] = self.df['Адрес доставки'].str.replace(pattern_postal_code_region, '', regex=True)

        # More transformations and cleaning
        self.extract_detailed_location_info()

    def extract_detailed_location_info(self):
        # Regex patterns
        pattern_postal_code_region = r'^\d+,\s[^\,]+,\s+'
        pattern_city = r'^(?:[А-Яа-я]+\s+)([^\,]+)'
        pattern_locality = (
            r'мкр (\d+)'                          # Match "мкр" followed by one or more digits
            r'|ул им [А-Я\.]+\s*([А-Я][а-я]+)'    # Match "ул им" followed by uppercase letters/periods, optional spaces, and a word starting with an uppercase followed by lowercase letters
            r'|ул [А-Я.]+\b\s*([\w\s\-]+)'        # Match "ул" followed by an uppercase letter, a period, and a word that may include alphanumeric characters, spaces, or hyphens
            r'|тер ([\w-]+)'                      # Match "тер" followed by one or more alphanumeric characters or hyphens
            r'|ул им ([а-яА-ЯёЁ-]+)'              # Match "ул им" followed by one or more characters from the Russian alphabet or hyphens
            r'|проезд (\d+-й)'                    # Match "проезд" follwed by one or more any symbols.
            r'|ул ([а-яА-ЯёЁ-]+)'                # Match just only simple one-word street title.
            r'|мкр (Сити-3)'
        )

        pattern_smallloc = r'[А-Яа-я\s]+,([А-Яа-я\s]),\w+'
        pattern_building = r'д (\d+\s*[кК]?\s*\d*|\d+[А-я]?)'
        pattern_apart = r'кв (\d+)'


        # Remove the postal code and region
        self.df['Адрес доставки 2'] = self.df['Адрес доставки'].str.replace(pattern_postal_code_region, '', regex=True)

        # Extract city, locality, building, unit and apartment numbers more accurately
        self.df['City'] = self.df['Адрес доставки 2'].str.extract(pattern_city)
        self.df['Building Number'] = self.df['Адрес доставки 2'].str.extract(pattern_building)[0]
        self.df['Building Number'] = self.df['Building Number'].str.replace(' ', '').str.upper()
        self.df['Apartment Number'] = self.df['Адрес доставки 2'].str.extract(pattern_apart)
        locality_extract = self.df['Адрес доставки 2'].str.extract(pattern_locality)
        self.df['Locality'] = locality_extract.apply(
            lambda x: f"{x[0]}-й микрорайон" if pd.notna(x[0]) else
                      ("ул " + x[1].strip() if pd.notna(x[1]) else
                       ("ул " + x[2].strip() if pd.notna(x[2]) else
                        (f"{x[3]}" if pd.notna(x[3]) else
                         (f"ул {x[4].strip()}" if pd.notna(x[4]) else
                          (f"{x[5].strip()} проезд" if pd.notna(x[5]) else
                           (f"ул " + x[6].strip() if pd.notna(x[6]) else 
                            (f"{x[7]}" if pd.notna(x[7]) else ""))))))),
            axis=1
        )

        self.df['SmallLoc'] = self.df['Адрес доставки 2'].str.split(',').str[1].str.strip()
        self.df['Адрес доставки 2'] = self.df['Адрес доставки 2'].str.replace(r'г Элиста, ', '')


        # Strip strings in column and replace many spaces
        self.df['Комментарий'] = self.df['Комментарий'].str.replace(r'\s+', ' ', regex=True).str.strip()
        self.df['Комментарий'] = self.df['Комментарий'].str.replace(r'комментарий:(?:\s+)', '', regex=True)
        self.df['Комментарий'] = self.df['Комментарий'].str.replace(r'домофон:(?:\s+)', 'д', regex=True)

    def group_and_aggregate_by_locality(self):
        # Initialize an empty DataFrame to hold the combined data
        all_areas_df = pd.DataFrame()
        # Define unique locality for filtering df
        unique_localities = self.df['SmallLoc'].unique()
        pattern = r'(\d+)'

        # Filter and aggregate data by locality
        for locality in unique_localities:
            df_area = self.df[self.df['Адрес доставки'].str.contains(rf"{locality}+\b", regex=True, na=False)].copy()
            if df_area.empty:
                continue

            df_area['Building Number 2'] = df_area['Building Number'].str.extract(pattern).astype(int)
            df_area = df_area.sort_values(by='Building Number 2')
            df_area['Кол-во посылок в партии'] = pd.to_numeric(df_area['Кол-во посылок в партии'], errors='coerce')

            summary_data = {
                'ФИО Получателя\xa0/Отправителя': 'ИТОГО:',
                'Адрес доставки 2': df_area['Адрес доставки'].count(),
                'Кол-во посылок в партии': df_area['Кол-во посылок в партии'].sum(),
                'Время доставки': len(df_area['Адрес доставки'].unique())
            }
            summary_row = pd.DataFrame(summary_data, index=[0])
            df_area = pd.concat([df_area, summary_row], ignore_index=True)

            all_areas_df = pd.concat([all_areas_df, df_area], ignore_index=True)

        all_areas_df.index = range(1, len(all_areas_df) + 1)
        new_order = ['ФИО Получателя\xa0/Отправителя',
       'Адрес доставки 2', 'Кол-во посылок в партии', 'Телефон', 'Время доставки',
       'Согласованная дата доставки', 'Комментарий',]
        self.all_areas_df = all_areas_df[new_order]
    
    def get_main_dataframe(self):
        return self.df
    
    def get_aggregated_dataframe(self):
        self.group_and_aggregate_by_locality()
        return self.all_areas_df

class ExcelManager:
    def __init__(self, df):
        self.df = df

    def save_with_format(self, filename, sheet_name):
        # Create a Pandas Excel writer using XlsxWriter as the engine
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        self.df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Get the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Set the header format
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1
        })
        font_bold = workbook.add_format({'bold': True})  # Bold format for specific usage

        # Write the column headers with the defined format
        for col_num, value in enumerate(self.df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        # Set column widths and row formats
        self.set_column_widths(worksheet, font_bold)
        self.apply_conditional_formatting(worksheet, workbook)

        # Close the Pandas Excel writer and output the Excel file
        writer.close()

    def set_column_widths(self, worksheet, font_bold):
        # Adjust column widths
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:B', 35)
        worksheet.set_column('C:C', 4, font_bold)  # Bold format for column C
        worksheet.set_column('D:F', 15)
        worksheet.set_column('G:G', 30)
        worksheet.set_row(0, 30)  # Header row height

    def apply_conditional_formatting(self, worksheet, workbook):
        # Apply conditional formatting and styles to the worksheet
        nepravilno_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        itog_row_format = workbook.add_format({'bg_color': '#95f0aa', 'bold': True})
        summarno_format = workbook.add_format({'bg_color': '#32d858', 'bold': True})
        font_bold = workbook.add_format({'bold': True})


        # Set specific row formats, assuming identification of rows is handled elsewhere
        # for idx, row in df.iterrows():
        #     if row['SomeColumn'] == 'SomeValue':
        #         worksheet.set_row(idx + 1, None, itog_row_format)
        # Apply a conditional format to specific row
        for idx, row in self.df.iterrows():
            if row['ФИО Получателя\xa0/Отправителя'] == 'ИТОГО:':
                worksheet.set_row(idx, None, itog_row_format)


class Geocoder:
    def __init__(self):
        self.geolocator = Nominatim(user_agent="example_app_yourappname")
        # Using RateLimiter to handle the request delay
        self.geocode = RateLimiter(self.geolocator.geocode, min_delay_seconds=1)

    def geocode_address(self, row):
        try:
            address = f"{row['Building Number']}, {row['Locality']}, {row['City']}"
            location = self.geocode(address)
            return (location.latitude, location.longitude) if location else (None, None)
        except Exception as e:
            print(f"Error geocoding address: {address} - {e}")
            return (None, None)

    def apply_geocoding(self, df):
        df['Coordinates'] = df.apply(self.geocode_address, axis=1)

        
class MapVisualizer:
    def __init__(self):
        self.map = None

    def create_map(self, df, start_location=None, zoom_start=12):
        # Determine the center of the map
        if start_location is None:
            # Calculate average latitude and longitude only from valid coordinates
            valid_coords = df['Coordinates'].dropna()  # Remove None entries
            valid_coords = [coord for coord in valid_coords if coord[0] is not None and coord[1] is not None]
            avg_latitude = pd.Series([coord[0] for coord in valid_coords]).mean()
            avg_longitude = pd.Series([coord[1] for coord in valid_coords]).mean()

        # Create a Folium map centered around the average coordinates
        self.map = folium.Map(location=[avg_latitude, avg_longitude], zoom_start=zoom_start)
        marker_cluster = MarkerCluster().add_to(self.map)

        # Iterate over DataFrame rows
        for idx, row in df.iterrows():
            # Extract coordinates
            coords = row['Coordinates']
            # Check if coordinates are valid (not None and containing two items)
            if coords and len(coords) == 2 and all(coord is not None for coord in coords):
                folium.Marker(
                    location=[coords[0], coords[1]],
                    popup=(f"<strong>Адрес:</strong> {row['Адрес доставки 2']}<br>"
                           f"<strong>Телефон:</strong> {row['Телефон']}<br>"
                           f"<strong>Кол-во:</strong> {row['Кол-во посылок в партии']}<br>"
                           #f"<strong>№ посылки:</strong> {row['№ посылки']}<br>"
                           f"<label><input type='checkbox' class='marker-checkbox' data-marker-id='marker_{idx}'> \
                           Доставлено</label>"),
                    icon=folium.Icon(icon='home', color='red')
                ).add_to(marker_cluster)

        return self.map
    
    def save_map(self, file_path='map.html'):
        # Save the map to an HTML file
        if self.map is not None:
            self.map.save(file_path)
        else:
            raise ValueError("Map has not been created yet.")


class JavaScriptInjector:
    def __init__(self, file_path):
        self.file_path = file_path
        self.content = self._read_html_file()

    def _read_html_file(self):
        # Read the HTML file content
        with open(self.file_path, 'r', encoding='utf-8') as file:
            return file.read()

    def extract_marker_ids(self):
        # Parse the HTML content using BeautifulSoup
        soup = BeautifulSoup(self.content, 'html.parser')

        # Find all script tags in the HTML
        script_tags = soup.find_all('script')

        # Pattern to find marker declarations (assuming var declarations)
        pattern = re.compile(r"var\s+(marker_\w+)\s+=\s+L.marker\b")

        # Set to store unique marker IDs
        marker_ids = set()

        # Extract and store all marker IDs found in script tags
        for script in script_tags:
            if script.string:  # Check if the script tag contains anything
                matches = pattern.findall(script.string)
                if matches:
                    marker_ids.update(matches)

        return marker_ids

    def generate_js_code(self, marker_ids):
        js_code = ""
        for marker_id in marker_ids:
            js_code += f"initializePopupListeners({marker_id});\n"
        return js_code

    def inject_javascript(self):
        marker_ids = self.extract_marker_ids()
        js_code_initialize_popup_listeners = self.generate_js_code(marker_ids)

        # JavaScript to add for handling the checkbox interaction and changing marker color
        js_code = f"""
        <script>
        // Function to initialize event listeners on popup open
            function initializePopupListeners(marker) {{
                marker.getPopup().on("add", function () {{
                    var checkbox = this.getContent().querySelector("input.marker-checkbox");
                    if (checkbox) {{
                        checkbox.addEventListener("change", function () {{
                            if (this.checked) {{
                                marker.setIcon(
                                    L.AwesomeMarkers.icon({{
                                        icon: "home",
                                        markerColor: "green", // Color when checked
                                        prefix: "glyphicon",
                                        iconColor: "white",
                                    }}),
                                );
                            }} else {{
                                marker.setIcon(
                                    L.AwesomeMarkers.icon({{
                                        icon: "home",
                                        markerColor: "red", // Color when unchecked
                                        prefix: "glyphicon",
                                        iconColor: "white",
                                    }}),
                                );
                            }}
                        }});
                    }}
                }});
            }}

            // Initialize popup listeners for each marker
            {js_code_initialize_popup_listeners}
            // Add similar lines for each marker you have

        </script>
        """

        # Insert the script after the last </script> tag
        last_script_pos = self.content.rfind('</script>')
        if last_script_pos != -1:
            # Insert the new script after the last </script>
            self.content = self.content[:last_script_pos + 9] + js_code + self.content[last_script_pos + 9:]
        else:
            # If no </script> tag is found, append at the end of the file
            self.content += js_code

    def save_modified_html(self, output_file_path=None):
        # Save the modified HTML content to a file
        if output_file_path is None:
            output_file_path = self.file_path
        with open(output_file_path, 'w', encoding='utf-8') as file:
            file.write(self.content)


def main():
    input_filename = 'Маршрутный лист.html'
    processor = DataProcessor(input_filename)
    processor.read_and_parse_html()
    processor.clean_and_transform_data()
    processor.extract_detailed_location_info()

    excel_manager = ExcelManager(processor.get_aggregated_dataframe())
    output_filename = 'Маршрутный лист.xlsx'
    excel_manager.save_with_format(output_filename, 'Маршрут')

    geocoder = Geocoder()
    geocoder.apply_geocoding(processor.get_main_dataframe())

    map_visualizer = MapVisualizer()
    map_created = map_visualizer.create_map(processor.get_main_dataframe())
    map_visualizer.save_map('output_map.html')

    injector = JavaScriptInjector('output_map.html')
    injector.inject_javascript()
    injector.save_modified_html('modified_map.html')

if __name__ == "__main__":
   main()
