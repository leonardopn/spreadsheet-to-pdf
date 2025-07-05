import pandas as pd
from fpdf import FPDF
import os


class PDF(FPDF):
    """
    Classe personalizada para criar o PDF com cabeçalho e rodapé.
    """

    def header(self):
        # Define a fonte para o cabeçalho
        self.set_font("Arial", "B", 12)
        # Título
        self.cell(0, 10, "Relatório de Dados do Excel", 0, 1, "C")
        # Pula uma linha
        self.ln(10)

    def footer(self):
        # Posiciona o cursor a 1.5 cm do final da página
        self.set_y(-15)
        # Define a fonte para o rodapé
        self.set_font("Arial", "I", 8)
        # Número da página
        self.cell(0, 10, f"Página {self.page_no()}/{{nb}}", 0, 0, "C")


def dataframe_para_pdf(df, nome_arquivo_pdf):
    """
    Converte um DataFrame do pandas para um arquivo PDF formatado.

    :param df: O DataFrame a ser convertido.
    :param nome_arquivo_pdf: O nome do arquivo PDF de saída.
    """
    if df is None:
        print("DataFrame não está disponível. A geração do PDF foi cancelada.")
        return

    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=10)

    # --- Largura das Colunas (Cálculo dinâmico) ---
    df_str = df.astype(str)
    larguras_colunas = []
    for col in df_str.columns:
        largura_cabecalho = pdf.get_string_width(col) + 6
        largura_max_dados = (
            df_str[col].apply(lambda x: pdf.get_string_width(x)).max() + 6
        )
        larguras_colunas.append(max(largura_cabecalho, largura_max_dados))

    # --- Cabeçalho da Tabela ---
    pdf.set_fill_color(200, 220, 255)
    for i, header in enumerate(df.columns):
        pdf.cell(
            larguras_colunas[i], 10, text=str(header), border=1, align="C", fill=True
        )
    pdf.ln()

    # --- Dados da Tabela (com linhas zebradas) ---
    pdf.set_fill_color(255, 255, 255)
    fill = False
    for index, row in df.iterrows():
        for i, data in enumerate(row):
            pdf.cell(
                larguras_colunas[i], 10, text=str(data), border=1, align="L", fill=fill
            )
        pdf.ln()
        fill = not fill

    pdf.output(nome_arquivo_pdf)
    print(f"\nPDF '{nome_arquivo_pdf}' gerado com sucesso!")


def main():
    """
    Função principal que orquestra a execução do script.
    """

    # --- Passo 2: Preparar os dados do Excel ---
    nome_arquivo_excel = "test.xlsx"
    if not os.path.exists(nome_arquivo_excel):
        print(
            f"Arquivo '{nome_arquivo_excel}' não encontrado. Criando um arquivo de exemplo."
        )
        dados_exemplo = {
            "Produto": [
                "Notebook Dell",
                "Mouse sem Fio",
                "Teclado Mecânico",
                "Monitor 4K",
            ],
            "Preço (R$)": [4500.50, 120.75, 299.90, 1800.00],
            "Quantidade em Estoque": [15, 80, 45, 22],
            "Data de Cadastro": [
                "2023-01-10",
                "2023-02-15",
                "2023-03-20",
                "2023-04-25",
            ],
        }
        df_exemplo = pd.DataFrame(dados_exemplo)
        df_exemplo.to_excel(nome_arquivo_excel, index=False)

    # --- Passo 3: Carregar os dados ---
    try:
        df = pd.read_excel(nome_arquivo_excel)
        print("Arquivo Excel carregado com sucesso!")
        print("\nVisualização dos 5 primeiros registros:")
        print(df.head())
    except FileNotFoundError:
        print(
            f"Erro: Arquivo '{nome_arquivo_excel}' não encontrado. Verifique o nome e o local do arquivo."
        )
        return  # Encerra a execução se o arquivo não puder ser lido

    # --- Passo 4: Gerar o PDF ---
    nome_do_pdf = "tabela_convertida.pdf"
    dataframe_para_pdf(df, nome_do_pdf)


if __name__ == "__main__":
    main()
