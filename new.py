import msoffcrypto
import pandas as pd
import io

def load_secure_mapping(file_path, password):
    decrypted_memory_stream = io.BytesIO()
    
    # 1. Open the encrypted file and decrypt it directly into RAM
    with open(file_path, 'rb') as file:
        office_file = msoffcrypto.OfficeFile(file)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_memory_stream)
    
    # 2. Read the RAM stream into pandas
    # using usecols="B:C" to only grab your PX and RM columns
    df = pd.read_excel(decrypted_memory_stream, engine='openpyxl', usecols="B:C")
    
    # 3. Rename columns to be safe, assuming row 1 might be a header or empty
    df.columns = ['PX_Code', 'RM_Code']
    
    # Drop any empty rows if your table has gaps
    df = df.dropna()
    
    # 4. Convert the two columns into a super-fast Python dictionary
    # Format: {'PX001': 'RM032', 'PX023': 'RM032', ...}
    mapping_dict = dict(zip(df['PX_Code'], df['RM_Code']))
    
    return mapping_dict

# --- How your Watchdog uses it ---

# RUN THIS ONCE when the watchdog script starts:
print("Unlocking mapping table into memory...")
SECURE_MAPPING = load_secure_mapping('test.xlsx', 'password')
print("Mapping loaded and secured in RAM.")

def process_incoming_file(incoming_px):
    # Instant, O(1) lookup. No file opening required.
    # The .get() method returns a default message if the PX code isn't found.
    rm_code = SECURE_MAPPING.get(incoming_px, "ERROR: PX code not found")
    return rm_code

# Example watchdog trigger:
result = process_incoming_file('PX023')
print(f"Result for workflow: {result}")