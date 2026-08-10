import socket
import threading
from datetime import datetime

# Server Configuration
# '0.0.0.0' listens on all active IPs, including 192.168.4.1
HOST = '0.0.0.0'
PORT = 80          # Set this to the target TCP port your Android app connects to
BUFFER_SIZE = 4096

def handle_client(client_socket, client_address):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"\n[{timestamp}] 🟢 Incoming TCP Connection from: {client_address[0]}:{client_address[1]}")
    
    try:
        # Receive incoming TCP payload from Android app
        data = client_socket.recv(BUFFER_SIZE)
        
        if data:
            print("=" * 60)
            print("📥 RAW TCP PAYLOAD (HEX):")
            print(data.hex(' '))
            print("-" * 60)
            print("📥 RECEIVED DATA (TEXT / STRING):")
            try:
                print(data.decode('utf-8'))
            except UnicodeDecodeError:
                print("[!] Data contains raw binary bytes, could not decode as UTF-8.")
            print("=" * 60)

            # Send a response back to the Android app
            # Option A: Standard HTTP 200 OK Response (if app expects HTTP)
            response = (
                "HTTP/1.1 200 OK\r\n"
                "Content-Type: text/plain\r\n"
                "Content-Length: 2\r\n"
                "Connection: close\r\n\r\n"
                "OK"
            )
            
            # Option B: Raw TCP Acknowledgement (uncomment if app expects plain text/bytes instead of HTTP)
            # response = "OK\n"
            
            client_socket.sendall(response.encode('utf-8'))
            print("📤 Sent response back to Android app.")
            
    except Exception as e:
        print(f"❌ Error handling TCP client: {e}")
    finally:
        client_socket.close()
        print(f"🔴 Connection closed with {client_address[0]}:{client_address[1]}\n")

def start_tcp_server():
    server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    # Allows reusing the port immediately after restart
    server_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)

    try:
        server_socket.bind((HOST, PORT))
        server_socket.listen(5)
        print("*" * 60)
        print(f"🚀 TCP/IP Server Running!")
        print(f"📡 Listening on Port: {PORT}")
        print("Waiting for incoming Android app connections...")
        print("*" * 60)

        while True:
            client_socket, client_address = server_socket.accept()
            # Handle client in a new thread
            client_thread = threading.Thread(
                target=handle_client, 
                args=(client_socket, client_address)
            )
            client_thread.start()

    except KeyboardInterrupt:
        print("\n🛑 Shutting down TCP server...")
    except Exception as e:
        print(f"\n❌ Socket Error: {e}")
    finally:
        server_socket.close()

if __name__ == "__main__":
    start_tcp_server()