import pandas as pd
from tkinter import filedialog
from customtkinter import *
from tkcalendar import DateEntry
import win32com.client as win32

# Instanciar a nossa tela 
tela = CTk()

# Configuração da tela
tela.title("Automation Invite")
tela.geometry("582x600")
tela.resizable(False, False)
tela.config(background="#6A737B")

# Criar a Frame da primeira página
frame_primeira_pagina = CTkFrame(tela, width=500, height=550)
frame_primeira_pagina.place(relx=0.5, rely=0.5, anchor=CENTER)

# Adicionar label com o nome do app
label_nomeApp = CTkLabel(frame_primeira_pagina, text="Automation Invite", bg_color="transparent", font=("Arial", 25, "bold"))
label_nomeApp.place(relx=0.50, rely=0.10, anchor=CENTER)

# Informações para a Textbox
info = """
NOTA: Este algoritmo foi desenvolvido para a automatização do envio de convites "Lembrete" em massa. Utilize sempre uma base de dados EXCEL (.xlsx) para esta automação.

PASSO A PASSO:
1. Crie uma base de dados .xlsx e adicione uma tabela com os campos "Hora" e "Email". SALVE O ARQUIVO NO C:
NOTA: Os caracteres "H" e "E" serão em caixa alta.
"""

# Adicionar Textbox com informações
textbox_info = CTkTextbox(frame_primeira_pagina, width=490, height=200, bg_color="transparent", fg_color="#5b646b", scrollbar_button_color="#48494a")
textbox_info.insert("0.0", info)
textbox_info.configure(state="disabled")  # Desabilitar edição
textbox_info.place(relx=0.50, rely=0.45, anchor=CENTER)

def nav_tela2():
    # Esquecer a tela 1
    frame_primeira_pagina.forget()

    # Criar a tela 2
    frame_segunda_pagina = CTkFrame(tela, width=500, height=550)
    frame_segunda_pagina.place(relx=0.5, rely=0.5, anchor=CENTER)
    
    def upload_arqv():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df = pd.read_excel(file_path)
            textbox_excel.configure(state="normal")
            textbox_excel.delete("1.0", "end")
            textbox_excel.insert("0.0", df.to_string())
            textbox_excel.configure(state="disabled")

            def nav_tela3():
                frame_segunda_pagina.forget()

                # Criar a tela 3
                frame_terceira_pagina = CTkFrame(tela, width=500, height=550)
                frame_terceira_pagina.place(relx=0.5, rely=0.5, anchor=CENTER)

                # Label e Entry para o assunto do e-mail
                label_assunto = CTkLabel(frame_terceira_pagina, text="Assunto:", font=("Arial", 10), bg_color="transparent")
                label_assunto.place(relx=0.10, rely=0.05)
                entry_assunto = CTkEntry(frame_terceira_pagina, width=400)
                entry_assunto.place(relx=0.10, rely=0.10)

                # Label e Textbox para o corpo do e-mail
                label_corpo_email = CTkLabel(frame_terceira_pagina, text="Corpo E-mail:", font=("Arial", 10), bg_color="transparent")
                label_corpo_email.place(relx=0.10, rely=0.16)
                textbox_corpo_email = CTkTextbox(frame_terceira_pagina, width=400, height=100)
                textbox_corpo_email.place(relx=0.10, rely=0.21)

                # Label e DateEntry para escolher a data da reunião
                label_data_reuniao = CTkLabel(frame_terceira_pagina, text="Data da Reunião:", font=("Arial", 10), bg_color="transparent")
                label_data_reuniao.place(relx=0.10, rely=0.40)
                data_reuniao_entry = DateEntry(frame_terceira_pagina, width=12, background='darkblue', foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
                data_reuniao_entry.place(relx=0.10, rely=0.45)
                data_reuniao_entry.config(state='readonly')  # Desabilitar edição

                # Label e Combobox para escolher a duração da reunião
                label_duracao_reuniao = CTkLabel(frame_terceira_pagina, text="Duração da Reunião:", font=("Arial", 10), bg_color="transparent")
                label_duracao_reuniao.place(relx=0.10, rely=0.50)
                combo_duracao = CTkComboBox(frame_terceira_pagina, values=["10 minutos", "15 minutos", "20 minutos"], width=400)
                combo_duracao.set("10 minutos")  # Define o valor padrão
                combo_duracao.place(relx=0.10, rely=0.55)
                combo_duracao.configure(state="readonly")  # Desabilitar edição

                # Label e Combobox para escolher o lembrete da reunião
                label_lembrete = CTkLabel(frame_terceira_pagina, text="Lembrete:", font=("Arial", 10), bg_color="transparent")
                label_lembrete.place(relx=0.1, rely=0.61)
                combo_lembrete = CTkComboBox(frame_terceira_pagina, values=["5 minutos", "10 minutos", "15 minutos"], width=400)
                combo_lembrete.set("5 minutos")  # Define o valor padrão
                combo_lembrete.place(relx=0.1, rely=0.66)
                combo_lembrete.configure(state="readonly")  # Desabilitar edição

                # Label e Entry para o local da reunião
                label_local_reuniao = CTkLabel(frame_terceira_pagina, text="Local:", font=("Arial", 10), bg_color="transparent")
                label_local_reuniao.place(relx=0.1, rely=0.72)
                entry_local_reuniao = CTkEntry(frame_terceira_pagina, width=400)
                entry_local_reuniao.place(relx=0.1, rely=0.77)

                # Botão para agendar 
                def agendar():
                    assunto = entry_assunto.get()
                    corpo_email = textbox_corpo_email.get("1.0", "end-1c")
                    data = data_reuniao_entry.get()
                    duracao = int(combo_duracao.get().split()[0])
                    lembrete = int(combo_lembrete.get().split()[0])
                    local = entry_local_reuniao.get()

                    # Agendar as reuniões
                    for index, linha in df.iterrows():
                        hora = linha['Hora']
                        email = linha['Email']
                        if pd.notna(hora) and isinstance(email, str):
                            agendar_reuniao(email, assunto, corpo_email, data, hora, duracao, lembrete, local)

                    # Esquecer a tela 3
                    frame_terceira_pagina.forget()

                    # Criar a tela 4 (Quarta Página) com a lista de reuniões agendadas
                    frame_quarta_pagina = CTkFrame(tela, width=500, height=550)
                    frame_quarta_pagina.place(relx=0.5, rely=0.5, anchor=CENTER)

                    # Label para mostrar lista de reuniões agendadas
                    label_lista_agendadas = CTkLabel(frame_quarta_pagina, text="Reuniões Agendadas", font=("Arial", 14, "bold"), bg_color="transparent")
                    label_lista_agendadas.place(relx=0.5, rely=0.1, anchor=CENTER)

                    # Exemplo de lista de reuniões agendadas
                    for index, linha in df.iterrows():
                        hora = linha['Hora']
                        email = linha['Email']
                        if pd.notna(hora) and isinstance(email, str):
                            label_reuniao = CTkLabel(frame_quarta_pagina, text=f"Email: {email}, Data: {data}, Hora: {hora}", font=("Arial", 10), bg_color="transparent")
                            label_reuniao.place(relx=0.5, rely=0.2 + index*0.05, anchor=CENTER)
                    
                    # Função para simular carregamento
                def simular_carregamento():
                    # Criar uma barra de progresso
                    progress_bar = CTkProgressBar(frame_terceira_pagina, width=400, height=5)
                    progress_bar.place(relx=0.50, rely=0.95, anchor=CENTER)
                    progress_bar.set(0)

                    # Atualizar a barra de progresso
                    for i in range(1, 61):
                        tela.update_idletasks()
                        progress_bar.set(i / 60)
                        tela.after(5)  # Simular um atraso   


                    # Agendar reuniões
                    agendar()       

                botao_agendar = CTkButton(frame_terceira_pagina, text="Agendar", width=200, bg_color="transparent", command=simular_carregamento)
                botao_agendar.place(relx=0.71, rely=0.9, anchor=CENTER)
            
            # botão confirmar
            botao_confirmar = CTkButton(frame_segunda_pagina, command=nav_tela3, text="Confirmar", width=200, bg_color="transparent", fg_color="black")
            botao_confirmar.place(relx=0.79, rely=0.75, anchor=CENTER)

    # Label do botão
    label_botao = CTkLabel(frame_segunda_pagina, text="Clique no botão para carregar um arquivo", font=("arial", 9))
    label_botao.place(relx=0.50, rely=0.15, anchor=CENTER)

    # Botão de upload de arquivo
    botao_upload = CTkButton(frame_segunda_pagina, text="Upload Excel File", command=upload_arqv, width=200, bg_color="transparent", fg_color="black")
    botao_upload.place(relx=0.50, rely=0.20, anchor=CENTER)

    # Adicionar Textbox para exibir o conteúdo do arquivo Excel
    textbox_excel = CTkTextbox(frame_segunda_pagina, width=490, height=200, bg_color="transparent", fg_color="#5b646b", scrollbar_button_color="#48494a")
    textbox_excel.configure(state="disabled")  # Desabilitar edição
    textbox_excel.place(relx=0.50, rely=0.45, anchor=CENTER)

# Adicionar botão para avançar na primeira tela
botao_avancar = CTkButton(frame_primeira_pagina, command=nav_tela2, text="Avançar", width=200, bg_color="transparent", fg_color="black")
botao_avancar.place(relx=0.5, rely=0.75, anchor=CENTER)

def agendar_reuniao(email, assunto, corpo_email, data, hora, duracao, lembrete, local):
    outlook = win32.Dispatch("Outlook.Application")
    appt = outlook.CreateItem(1)  
    appt.Start = f"{data} {hora}"
    appt.Subject = assunto
    appt.Duration = duracao
    appt.Location = local
    appt.MeetingStatus = 1  
    appt.Body = corpo_email
    appt.ReminderMinutesBeforeStart = lembrete
    appt.Recipients.Add(email)
    appt.Save()
    appt.Send()

# Run the application
tela.mainloop()
