import os
import base64
import json
import csv
import pandas as pd
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
from threading import Thread

# Configuración básica
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

class EnviadorRecibos:
    def __init__(self, log_callback=None):
        self.creds = None
        self.service = None
        self.log_callback = log_callback
        self._autenticar()

    def _log(self, message):
        if self.log_callback:
            self.log_callback(message)

    def _autenticar(self):
        if os.path.exists('token.json'):
            try:
                with open('token.json', 'r') as token_file:
                    self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            except Exception as e:
                self._log(f"Error leyendo token: {str(e)}")
                os.remove('token.json')

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
                self.creds = flow.run_local_server(port=0)

            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())

        self.service = build('gmail', 'v1', credentials=self.creds)
        self._log("✅ Autenticación exitosa")

    def _buscar_recibos(self, dni):
        if not os.path.exists('recibos'):
            os.makedirs('recibos')
        return [f for f in os.listdir('recibos') if dni in f and f.lower().endswith('.pdf')]

    def enviar_recibo(self, dni, email, modo_prueba=False):
        archivos = self._buscar_recibos(dni)
        if not archivos:
            return False, "No se encontraron recibos"

        msg = MIMEMultipart()
        msg['To'] = email
        msg['Subject'] = f"Recibo de Sueldo {datetime.now().strftime('%m/%Y')}"

        cuerpo = f"Estimado/a,\n\nAdjunto su recibo de sueldo (DNI: {dni}).\n\nSaludos"
        msg.attach(MIMEText(cuerpo))

        for archivo in archivos:
            with open(os.path.join('recibos', archivo), 'rb') as f:
                part = MIMEApplication(f.read(), Name=archivo)
                part['Content-Disposition'] = f'attachment; filename="{archivo}"'
                msg.attach(part)

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()

        if not modo_prueba:
            try:
                self.service.users().messages().send(
                    userId='me', body={'raw': raw}).execute()
                return True, "Enviado correctamente"
            except Exception as e:
                return False, str(e)
        return True, "Modo prueba - No enviado"

class InterfazApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Enviador de Recibos")
        self.root.geometry("600x400")
        
        self.modo_prueba = tk.BooleanVar(value=False)
        
        # Crear interfaz
        ttk.Label(root, text="Enviador de Recibos", font=('Arial', 14)).pack(pady=10)
        
        ttk.Checkbutton(root, text="Modo Prueba", variable=self.modo_prueba).pack()
        
        ttk.Button(root, text="Enviar Recibos", command=self.iniciar_envio).pack(pady=10, fill=tk.X)
        
        ttk.Button(root, text="Agregar Recibos", command=self.agregar_recibos).pack(pady=5, fill=tk.X)
        
        self.log = scrolledtext.ScrolledText(root, height=10)
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.enviador = EnviadorRecibos(log_callback=self.actualizar_log)
    
    def actualizar_log(self, mensaje):
        self.log.insert(tk.END, mensaje + "\n")
        self.log.see(tk.END)
        self.root.update()
    
    def iniciar_envio(self):
        Thread(target=self._enviar_todos, daemon=True).start()
    
    def _enviar_todos(self):
        if not os.path.exists('destinatarios.xlsx'):
            messagebox.showerror("Error", "No existe el archivo destinatarios.xlsx")
            return
        
        try:
            df = pd.read_excel('destinatarios.xlsx')
            for _, row in df.iterrows():
                dni, email = str(row['dni']), str(row['email'])
                success, msg = self.enviador.enviar_recibo(dni, email, self.modo_prueba.get())
                status = "✅" if success else "❌"
                self.actualizar_log(f"{status} {dni} - {email}: {msg}")
        except Exception as e:
            self.actualizar_log(f"❌ Error: {str(e)}")
    
    def agregar_recibos(self):
        archivos = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        if not archivos:
            return
        
        if not os.path.exists('recibos'):
            os.makedirs('recibos')
        
        for archivo in archivos:
            try:
                destino = os.path.join('recibos', os.path.basename(archivo))
                with open(archivo, 'rb') as f_origen, open(destino, 'wb') as f_destino:
                    f_destino.write(f_origen.read())
                self.actualizar_log(f"📄 Copiado: {os.path.basename(archivo)}")
            except Exception as e:
                self.actualizar_log(f"❌ Error copiando {archivo}: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = InterfazApp(root)
    root.mainloop()