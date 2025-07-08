import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import queue
import serial
import serial.tools.list_ports
import telnetlib3
import asyncio
import os.path

class Adress:
    
    # puxa o caminho %appdata%\config.txt
    config_path = os.environ.get('AppData')+r"\config.cng"
    
    # se o aruqivo não existe, cria um arquivo novo
    if not os.path.exists(config_path):
        arquivo = open(config_path, 'w+')
        arquivo.writelines(u'COM1\n9600\nHostname\n23\nFalse\n1\n8\nNone')
        arquivo.close()
        with open(config_path) as file:
            lines = file.read().splitlines()

    # se o arquivo existe apenas lê as informações internas
    else:
        with open(config_path) as file:
            lines = file.read().splitlines()

class ToolTip:
    def __init__(self, widget, msg, delay=1.5, follow=True):
        self.widget = widget
        self.msg = msg
        self.delay = int(delay * 1000)  # Em milisegundos
        self.follow = follow
        self.tip_window = None
        self.id = None
        self.x = self.y = 0
        
        # Eventos do Tolltip
        widget.bind('<Enter>', self.schedule)
        widget.bind('<Leave>', self.hidetip)
        widget.bind('<Motion>', self.mousemove)
    
    def schedule(self, event=None):
        self.unschedule()
        self.id = self.widget.after(self.delay, self.showtip)
    
    def unschedule(self):
        if self.id:
            self.widget.after_cancel(self.id)
            self.id = None
    
    def showtip(self):
        # Cria janela Tooltip
        self.tip_window = tk.Toplevel(self.widget)
        self.tip_window.wm_overrideredirect(True)  # remove a parte superior do tooltip
        
        # posição do tooltip
        x, y = self.calculate_position()
        self.tip_window.wm_geometry(f"+{x}+{y}")
        
        # adiciona as informações a janela
        label = tk.Label(self.tip_window, pady=5, padx=5, text=self.msg, justify='left',
                         background="#f5f5f5", relief='solid', borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack()
    
    def hidetip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None
        self.unschedule()
    
    def mousemove(self, event):
        if self.follow and self.tip_window:
            self.x = event.x_root + 15
            self.y = event.y_root + 15
            self.tip_window.wm_geometry(f"+{self.x}+{self.y}")
    
    def calculate_position(self):
        # posição da janela referente ao cursor
        self.x = self.widget.winfo_pointerx() + 15
        self.y = self.widget.winfo_pointery() + 15
        return self.x, self.y

class SerialTelnetBridge:
    
    # Interface do programa
    def __init__(self, root):
        self.root = root
        self.root.title("Ponte Serial-Telnet")
        self.root.geometry("485x650")
        
        # Frame Serial
        self.connection_frame_serial = ttk.LabelFrame(root, text="Configuração serial")
        self.connection_frame_serial.pack(padx=10, pady=5, fill=tk.X)
        
        # Seleção da porta serial
        ttk.Label(self.connection_frame_serial, text="Porta Serial:").grid(row=0, column=0, padx=(5,5), pady=(5,10), sticky='e')
        self.port_combo = ttk.Combobox(self.connection_frame_serial, width=10)
        self.port_combo.grid(row=0, column=1, pady=(5,10))
        ToolTip(self.port_combo, msg="Porta COM que será utilizada para a comunicação serial.", follow=True)
        
        # seleção do Baud Rate
        ttk.Label(self.connection_frame_serial, text="Baud Rate:").grid(row=0, column=2, padx=(5,5), pady=(5,10), sticky='e')
        self.baud_combo = ttk.Combobox(self.connection_frame_serial, width=10, values=[
            '4800', '9600', '19200', '38400', '57600', '115200', '230400', '460800'
        ])
        self.baud_combo.grid(row=0, column=3, pady=(5,10))
        self.baud_combo.set(Adress.lines[1])
        ToolTip(self.baud_combo, msg="Baud Rate da porta COM para a comunicação serial.", follow=True)
        
        # seleção do Stop Bits
        ttk.Label(self.connection_frame_serial, text="Stop Bits:").grid(row=1, column=0, padx=(5,5), pady=(5,10), sticky='e')
        self.stop_bits_combo = ttk.Combobox(self.connection_frame_serial, width=10, values=[
            '1', '1.5', '2'
        ])
        self.stop_bits_combo.grid(row=1, column=1, pady=(5,10))
        self.stop_bits_combo.set(Adress.lines[5])
        ToolTip(self.stop_bits_combo, msg="Stop Bits da porta COM para a comunicação serial.", follow=True)

        # seleção do Bist Size
        ttk.Label(self.connection_frame_serial, text="Bits Size:").grid(row=0, column=4, padx=(5,5), pady=(5,10), sticky='e')
        self.bits_size_combo = ttk.Combobox(self.connection_frame_serial, width=10, values=[
            '6', '7', '8'
        ])
        self.bits_size_combo.grid(row=0, column=5, pady=(5,10))
        self.bits_size_combo.set(Adress.lines[6])
        ToolTip(self.bits_size_combo, msg="Size Bits da porta COM para a comunicação serial.", follow=True)

        # seleção do Parity
        ttk.Label(self.connection_frame_serial, text="Parity:").grid(row=1, column=2, padx=(5,5), pady=(5,10), sticky='e')
        self.Parity_combo = ttk.Combobox(self.connection_frame_serial, width=10, values=[
            'None', 'Even', 'Odd', 'Mark', 'Space'
        ])
        self.Parity_combo.grid(row=1, column=3, pady=(5,10))
        self.Parity_combo.set(Adress.lines[7])
        ToolTip(self.Parity_combo, msg="Parity da porta COM para a comunicação serial.", follow=True)
        
        # Frame Telnet
        self.connection_frame_telnet = ttk.LabelFrame(root, text="Configuração Telnet(TCP/IP)")
        self.connection_frame_telnet.pack(padx=10, pady=5, fill=tk.X)

        # configuração Telnet
        ttk.Label(self.connection_frame_telnet, text="Telnet Host:").grid(row=0, column=0, sticky='w', pady=(10,10))
        self.host_entry = ttk.Entry(self.connection_frame_telnet, width=13)
        self.host_entry.grid(row=0, column=1, padx=5, pady=(10,10), sticky='w')
        self.host_entry.insert(0, Adress.lines[2])
        ToolTip(self.host_entry, msg="Host TCP para comunicação Telnet.\nExemplo: 192.168.1.1", follow=True)
        
        ttk.Label(self.connection_frame_telnet, text="Porta:").grid(row=0, column=2, padx=(15,0), pady=(10,10))
        self.port_entry = ttk.Entry(self.connection_frame_telnet, width=8)
        self.port_entry.grid(row=0, column=3, pady=(10,10), sticky='w')
        self.port_entry.insert(0, Adress.lines[3])
        ToolTip(self.port_entry, msg="Porta TCP para comunicação Telnet.\nExemplo: 23", follow=True)

        # Frame Telnet
        self.connection_frame_option = ttk.LabelFrame(root, text="Opções")
        self.connection_frame_option.pack(padx=10, pady=5, fill=tk.X)

        # botão de atualizar as portas seriais
        self.refresh_btn = ttk.Button(self.connection_frame_option, text="Atualizar Portas", command=self.refresh_ports)
        self.refresh_btn.grid(row=0, column=1, padx=10, pady=(10,0))
        ToolTip(self.refresh_btn, msg="Irá Atualizar a lista de portas COM existentes com base na maquina local.", follow=True)

        # botão de conexão
        self.connect_btn = ttk.Button(self.connection_frame_option, text="Conectar", command=self.toggle_connection)
        self.connect_btn.grid(row=0, column=0, padx=10, pady=(10,0))
        ToolTip(self.connect_btn, msg="Irá iniciar a ponte Telnet-Serial.", follow=True)

        # Botão de salvar configurações
        self.save_btn = ttk.Button(self.connection_frame_option, text="Salvar", command=self.save_config)
        self.save_btn.grid(row=0, column=2, padx=10, pady=(10,0))
        ToolTip(self.save_btn, msg="Utilizado para salvar as configurações definidas de Serial | Host-Telnet | Terminal.", follow=True)

        # Botão de esconder terminal
        self.hide_terminal_var = tk.BooleanVar(value=Adress.lines[4])
        self.hide_check = ttk.Checkbutton(
            self.connection_frame_option,
            text="Esconder terminal",
            variable=self.hide_terminal_var,
            command=self.hide_frame
        )
        self.hide_check.grid(row=1, column=0, columnspan=6, pady=0, sticky='w')
        ToolTip(self.hide_check, msg="Botão de checagem para caso queira ocultar o terminal de logs.", follow=True)

        # Terminal
        self.terminal_frame = ttk.LabelFrame(root, text="Terminal")
        self.terminal_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        if self.hide_terminal_var.get():
            self.terminal_frame.pack_forget()
            self.root.geometry("485x260")

        self.output_area = scrolledtext.ScrolledText(
            self.terminal_frame, 
            wrap=tk.WORD, 
            state='disabled'
        )
        self.output_area.pack(padx=5, pady=5, fill=tk.BOTH, expand=True)
        self.output_area.tag_config('serial', foreground='blue')
        self.output_area.tag_config('telnet', foreground='green')
        self.output_area.tag_config('error', foreground='red')
        self.output_area.tag_config('warning', foreground='purple')
        
        # variaveis iniciais
        self.connected = False
        self.serial_conn = None
        self.telnet_reader = None
        self.telnet_writer = None
        self.telnet_thread = None
        self.serial_thread = None
        self.data_queue = queue.Queue()
        self.loop = None
        self.refresh_ports()
        self.check_queue()
        
        # processo para fechar janela
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    # Função para fazer o terminal o ocultar/aparecer quando selecionado
    def hide_frame(self):
        if self.hide_terminal_var.get():
            self.terminal_frame.pack_forget()
            self.root.geometry("485x260")
        else:
            self.terminal_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
            self.root.geometry("485x650")

    # Puxa as portas seriais existentes no PC
    def refresh_ports(self):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.port_combo['values'] = ports
        if ports:
            self.port_combo.set(Adress.lines[0])

    # switch do botão conectar/desconectar
    def toggle_connection(self):
        if not self.connected:
            self.connect()
        else:
            self.disconnect()

    # função para botão desconectar
    def disconnect(self):
        if not self.connected:
            return

        self.connected = False
        self.data_queue.put(("warning", "Desconectando...\n"))
    
        # Fecha conexão serial
        if self.serial_conn and self.serial_conn.is_open:
            try:
                self.serial_conn.close()
            except Exception as e:
                self.data_queue.put(("error", f"Erro ao fechar serial: {str(e)}\n"))
    
        # Fecha conexão telnet
        if self.telnet_writer and not self.telnet_writer.transport.is_closing():
            try:
                async def close_writer():
                    await self.telnet_writer.close()
                asyncio.run_coroutine_threadsafe(close_writer(), self.loop)
            except Exception as e:
                self.data_queue.put(("error", f"Erro ao fechar telnet: {str(e)}\n"))
    
        # Para o loop asyncio
        if self.loop and self.loop.is_running():
            self.loop.call_soon_threadsafe(self.loop.stop)
    
        # Reseta o UI
        self.connect_btn.config(text="Conectar")
        self.set_ui_state(True)
        self.data_queue.put(("warning", "Desconectado. Pronto para nova conexão.\n"))

    # função para botão conectar
    def connect(self):
        # puxando parametros de conexão
        serial_port = self.port_combo.get()
        baud_rate = int(self.baud_combo.get())
        bits_size = int(self.bits_size_combo.get())
        stop_bits = float(self.stop_bits_combo.get())
        telnet_host = self.host_entry.get()
        telnet_port = int(self.port_entry.get())
        if self.Parity_combo.get()=='None':
            parity='N'
        elif self.Parity_combo.get()=='Even':
            parity='E'
        elif self.Parity_combo.get()=='Odd':
            parity='O'
        elif self.Parity_combo.get()=='Mark':
            parity='M'
        elif self.Parity_combo.get()=='Space':
            parity='S'
            
        if not serial_port:
            self.data_queue.put(("error", "Erro: Nenhuma porta serial selecionada!\n"))
            return

        # congela UI durante a tentativa de conexão
        self.set_ui_state(False)
        self.data_queue.put(("warning", f"Conectando á {serial_port}@{baud_rate} e {telnet_host}:{telnet_port}...\n"))

        try:
            # abre conexão serial
            self.serial_conn = serial.Serial(
                port=serial_port,
                baudrate=baud_rate,
                bytesize=bits_size,
                parity=parity,
                stopbits=stop_bits,
                timeout=1
            )
            
            # inicia o threads
            self.connected = True
            self.loop = asyncio.new_event_loop()
            
            self.telnet_thread = threading.Thread(
                target=self.start_telnet_client,
                args=(telnet_host, telnet_port),
                daemon=True
            )
            self.serial_thread = threading.Thread(
                target=self.serial_worker,
                daemon=True
            )
            
            self.telnet_thread.start()
            self.serial_thread.start()
            
            self.data_queue.put(("serial", f"Conectado! Lendo a serial: {serial_port}\n"))
            self.connect_btn.config(text="Desconectar")
            
        except Exception as e:
            self.data_queue.put(("error", f"erro de conexão: {str(e)}\n"))
            self.set_ui_state(True)
            if self.serial_conn and self.serial_conn.is_open:
                self.serial_conn.close()

    # Inicia o loop do asyncio para o telnetlib3
    def start_telnet_client(self, host, port):
        
        asyncio.set_event_loop(self.loop)
        self.loop.run_until_complete(self.telnet_worker(host, port))

    # Faz a comunicação com o Host telnet para fazer a leitura e mandar os dados
    async def telnet_worker(self, host, port):
        try:
            reader, writer = await telnetlib3.open_connection(host, port, shell=self.telnet_shell)
            self.telnet_reader = reader
            self.telnet_writer = writer
            self.data_queue.put(("telnet", f"Conectado á {host}:{port}\n"))
        
            # Mantem conexão viva
            while self.connected:
                await asyncio.sleep(0.1)
            
        except Exception as e:
            self.data_queue.put(("error", f"Telnet erro: {str(e)}\n"))
        finally:
            if self.connected:
                self.root.after(100, self.disconnect)
    
    # thread secundaria que segura os dados recebidos do telnet e redireciona para serial
    async def telnet_shell(self, reader, writer):
        while self.connected:
            try:
                # lê o telnet, com o valor entre () sendo o intervalo que ele faz essa operação
                data = await reader.read(100)
                if not data:
                    break
                    
                if self.serial_conn and self.serial_conn.is_open:
                    self.serial_conn.write(data.encode('ascii'))
                self.data_queue.put(("telnet", data+"\n"))
                
            except Exception as e:
                self.data_queue.put(("error", f"Erro no campo Telnet: {str(e)}\n"))
                break
    
    # thread secundaria para a leitura e escrita serial, e o redirecionamento da informação para o telnet
    def serial_worker(self):
        try:
            while self.connected and self.serial_conn and self.serial_conn.is_open:
                # lê a serial, com o valor entre () sendo o intervalo que ele faz essa operação
                data = self.serial_conn.read(100)
                if data:
                    decoded_data = data.decode('ascii', 'replace')
                    
                    # se conectado, manda para o telnet
                    if self.telnet_writer and not self.telnet_writer.transport.is_closing():
                        asyncio.run_coroutine_threadsafe(
                            self.telnet_writer.write(decoded_data),
                            self.loop
                        )
                    
                    # mostra informação no terminal
                    self.data_queue.put(("serial", decoded_data))
        except Exception as e:
            if self.connected:
                self.data_queue.put(("error", f"Erro na Serial: {str(e)}\n"))
                self.root.after(100, self.disconnect)
    
    # Checa novos dados para atualizar a UI
    def check_queue(self):
        try:
            while not self.data_queue.empty():
                data_type, data = self.data_queue.get_nowait()
                self.display_data(data_type, data)
        except queue.Empty:
            pass
        
        self.root.after(100, self.check_queue)
    
    # Formata os dados que chegam no terminal
    def display_data(self, data_type, data):
        
        self.output_area.config(state='normal')
        self.output_area.insert(tk.END, data, data_type)
        self.output_area.see(tk.END)
        self.output_area.config(state='disabled')

    # Ativa e desativa os controles UI
    def set_ui_state(self, enabled):
        state = 'normal' if enabled else 'disabled'
        self.port_combo.config(state=state)
        self.baud_combo.config(state=state)
        self.stop_bits_combo.config(state=state)
        self.bits_size_combo.config(state=state)
        self.Parity_combo.config(state=state)
        self.host_entry.config(state=state)
        self.port_entry.config(state=state)
        self.refresh_btn.config(state=state)
        self.save_btn.config(state=state)

    # puxa as informações selecionadas e atualiza o config.txt
    def save_config(self):
        serial_port = self.port_combo.get()
        baud_rate = self.baud_combo.get()
        telnet_host = self.host_entry.get()
        telnet_port = self.port_entry.get()
        checkbox = str(self.hide_terminal_var.get())
        stop_bits = self.stop_bits_combo.get()
        bits_size = self.bits_size_combo.get()
        parity = self.Parity_combo.get()
        arquivo = open(Adress.config_path, 'w+')
        arquivo.writelines(f"{serial_port}\n{baud_rate}\n{telnet_host}\n{telnet_port}\n{checkbox}\n{stop_bits}\n{bits_size}\n{parity}")
        arquivo.close()

    # função para desabilitar a conexão
    def on_closing(self):
        if self.connected:
            self.disconnect() 
            
        if self.loop and self.loop.is_running():
            self.loop.call_soon_threadsafe(self.loop.stop)
            
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialTelnetBridge(root)
    root.mainloop()