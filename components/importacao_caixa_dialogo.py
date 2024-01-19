"""COMPONENTE PARA CAIXA DE DIALOGO E INPUT DE USUARIO DO PYTHON"""
import tkinter as tk;

class DialogBox:
    def __init__(self, master):
        self.particao = None
        self.mes = None
        self.ano = None

        self.master = master
        self.master.title("Parâmetros Iniciais") # Título da Janela
        self.master.geometry("400x200")  # Tamanho da janela

        self.particao_var = tk.StringVar(self.master)
        self.entrada1_var = tk.StringVar(self.master)
        self.entrada2_var = tk.StringVar(self.master)

        self.error_message_label = tk.Label(master, text="", fg="red")
        self.error_message_label.pack()

        self.setup_ui()

    def setup_ui(self):
        # Título para o menu suspenso
        self.dropdown_title_label = tk.Label(self.master, text="Partição do Drive:")
        self.dropdown_title_label.pack()

        # Dropdown menu
        particoes = ["H", "I", "J", "K", "L", "M"]
        self.particao_var.set(particoes[0])  # Definir particao padrão
        menu_particoes = tk.OptionMenu(self.master, self.particao_var, *particoes)
        menu_particoes.pack()

        # Label e campo de entrada 1
        label1 = tk.Label(self.master, text="Mês")
        label1.pack()
        entrada1 = tk.Entry(self.master, textvariable=self.entrada1_var)
        entrada1.pack()

        # Label e campo de entrada 2
        label2 = tk.Label(self.master, text="Ano:")
        label2.pack()
        entrada2 = tk.Entry(self.master, textvariable=self.entrada2_var)
        entrada2.pack()

        # Botão de envio
        botao_submit = tk.Button(self.master, text="Enviar", command=self.on_submit)
        botao_submit.pack()

    def on_submit(self):
        # Ação a ser realizada quando o botão Submit for pressionado
        self.particao = self.particao_var.get()
        self.mes = self.entrada1_var.get()
        self.ano = self.entrada2_var.get()
        if (self.mes.isnumeric() == True  and int(self.mes) > 0 and int(self.mes) < 13) and self.ano.isnumeric() == True:
            if int(self.mes) < 10 and not self.mes.__contains__("0"):
                self.mes = "0" + self.mes
            self.error_message_label.config(text="")
            self.master.destroy()
        elif self.mes.isnumeric() == True and (int(self.mes) < 1 or int(self.mes) > 12):
            self.error_message_label.config(text="Por favor insira valores de 1 a 12 para o campo Mês.")
        elif self.mes.isnumeric() == False or self.ano.isnumeric() == False:
            self.error_message_label.config(text="Campos Mês e Ano so podem conter numeros inteiros! Tente Novamente.")
        else:
            self.error_message_label.config(text="Valores inválidos, por favor tente novamente.")
        


"""COLOCAR NO CODIGO PRINCIPAL
def main():
    root = tk.Tk()
    app = DialogBox(root)
    root.mainloop()
    return app.dropdown, app.mes, app.ano

if __name__ == "__main__":
    dropdown, mes, ano = main()
    print(f"Valor escolha: {dropdown}")
    print(f"Valor digitado 1: {mes}")
    print(f"Valor digitado 2: {ano}")
""" 