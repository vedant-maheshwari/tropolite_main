import firebase_admin
from firebase_admin import credentials, db
import time
import threading

cred = credentials.Certificate('excel-vault-331bd-firebase-adminsdk-fbsvc-41dd8e1c77.json')
firebase_admin.initialize_app(cred, {
    'databaseURL': 'https://excel-vault-331bd-default-rtdb.asia-southeast1.firebasedatabase.app'
})

# Signals the retry loop that the connection has dropped and we need to reconnect
connection_lost = threading.Event()

def converter(request_data, excel_password: str):
    items = request_data.get('items', [])
    print(items)
    print(f"[{time.strftime('%X')}] Excel password received: {excel_password!r}")
    # TODO: use excel_password here when opening the Excel file
    # e.g. open_excel_with_password(excel_password)
    return_list = []
    for i in range(len(items)):
        return_list.append({f'RM{items[i].get("Item_Code")[2:]}': items[i].get('Total_Qty')})
    return return_list

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
            vault = db.reference('vault_system').get()
            request = vault.get('request_data', {})
            excel_password = vault.get('excel_password', '')
            result = converter(request, excel_password)
            db.reference('vault_system').update({
                'status': 'completed',
                'final_result': result,
                'excel_password': None   # clear after use
            })
            print(f"[{time.strftime('%X')}] Matrix processed and results sent.")
        except Exception as e:
            print(f"Error processing: {e}")
            db.reference('vault_system').update({
                'status': 'error',
                'final_result': f"Laptop B Failed: {str(e)}"
            })

def start_listener_with_retry():
    """
    Start the Firebase listener. If no event is received within HEARTBEAT_TIMEOUT
    seconds, assume the connection dropped and reconnect automatically.
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
                print(f"[{time.strftime('%X')}] No heartbeat — connection likely dropped.")

        except Exception as e:
            print(f"[{time.strftime('%X')}] Listener error: {e}")

        finally:
            if listener:
                try:
                    listener.close()
                except Exception:
                    pass

        print(f"[{time.strftime('%X')}] Reconnecting in {retry_delay}s...")
        time.sleep(retry_delay)


if __name__ == '__main__':
    print('watchdog listening')

    listener_thread = threading.Thread(target=start_listener_with_retry, daemon=True)
    listener_thread.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("Shutting down Watchdog.")