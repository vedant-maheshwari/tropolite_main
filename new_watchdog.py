import firebase_admin
from firebase_admin import credentials, db
import time
import threading
import sys
import pandas as pd
import io
import msoffcrypto

cred = credentials.Certificate(
    'excel-vault-331bd-firebase-adminsdk-fbsvc-41dd8e1c77.json'
)

firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://excel-vault-331bd-default-rtdb.asia-southeast1.firebasedatabase.app'
})

# Signals the retry loop that the connection has dropped and we need to reconnect
connection_lost = threading.Event()


def converter(request_data, excel_passowrd):
    items = request_data.get('items', [])
    
    result_list = []

    try:
        # Load the encrypted Excel file into memory directly
        decrypted_memory_stream = io.BytesIO()
        with open("MDRM_1.xlsx", 'rb') as file:
            office_file = msoffcrypto.OfficeFile(file)
            office_file.load_key(password=excel_passowrd)
            office_file.decrypt(decrypted_memory_stream)
        
        # Read columns A and B into pandas
        df = pd.read_excel(decrypted_memory_stream, engine='openpyxl', usecols="A:B")
        df = df.dropna()
        
        # Build mapping dictionaries for fast, reliable lookup
        md_to_rm_mapping = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
        md_to_rm_mapping_str = {str(k).replace(".0", ""): v for k, v in md_to_rm_mapping.items()}

        for item in items:
            md_code = item.get("MD_code")
            
            if md_code in md_to_rm_mapping:
                result = md_to_rm_mapping[md_code]
                result_list.append({result: item.get("Qty")})
            elif str(md_code).replace(".0", "") in md_to_rm_mapping_str:
                result = md_to_rm_mapping_str[str(md_code).replace(".0", "")]
                result_list.append({result: item.get("Qty")})
            else:
                result_list.append({item.get('MD_code'): item.get('Qty')})
                
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        for item in items:
            result_list.append({item.get('MD_code'): item.get('Qty')})

    return result_list


def on_db_change(event):
    # Any event received means we're still connected — reset the lost flag
    connection_lost.clear()

    status = None

    if event.path == '/':
        if isinstance(event.data, dict):
            status = event.data.get('status')

    elif event.path == '/status':
        status = event.data

    if status == 'approved':
        print(f"[{time.strftime('%X')}] Approval received! converting...")

        try:
            request = db.reference('vault_system/request_data').get()
            vault = db.reference('vault_system').get()
            password = vault.get('excel_password', '')
            result = converter(request, password)

            db.reference('vault_system').update({
                'status': 'completed',
                'final_result': result
            })

            print(
                f"[{time.strftime('%X')}] Matrix processed and results sent."
            )

        except Exception as e:
            print(f"Error processing: {e}")

            db.reference('vault_system').update({
                'status': 'error',
                'final_result': f"Laptop B Failed: {str(e)}"
            })


def start_listener_with_retry():
    """
    Start the Firebase listener. If no event is received within
    HEARTBEAT_TIMEOUT seconds, assume the connection dropped and
    reconnect automatically.
    """
    HEARTBEAT_TIMEOUT = 150
    retry_delay = 5

    while True:
        listener = None
        connection_lost.clear()

        try:
            print(f"[{time.strftime('%X')}] Connecting to Firebase listener...")

            listener = db.reference('vault_system').listen(on_db_change)

            # Block until no event has been received for HEARTBEAT_TIMEOUT seconds.
            # connection_lost is set here to use as a "heartbeat missed" signal.
            timed_out = not connection_lost.wait(timeout=HEARTBEAT_TIMEOUT)

            if timed_out:
                print(
                    f"[{time.strftime('%X')}] No heartbeat — connection likely dropped."
                )

        except Exception as e:
            print(f"[{time.strftime('%X')}] Listener error: {e}")

        finally:
            if listener:
                try:
                    listener.close()
                except Exception:
                    pass

            print(
                f"[{time.strftime('%X')}] Reconnecting in {retry_delay}s..."
            )
            time.sleep(retry_delay)


if __name__ == '__main__':
    print('watchdog listening')

    listener_thread = threading.Thread(
        target=start_listener_with_retry,
        daemon=True
    )
    listener_thread.start()

    try:
        while True:
            time.sleep(1)

    except KeyboardInterrupt:
        print("Shutting down Watchdog.")
       