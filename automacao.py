import sys
import playwright
from playwright.sync_api import sync_playwright, TimeoutError
import openpyxl
import os
from pathlib import Path
import wget

def automacao(caminho_arquivo):

    with sync_playwright() as p:
        navegador = p.chromium.launch(headless=False)
        context = navegador.new_context()
        pagina = context.new_page()

        # Navegar até a página de login
        pagina.goto('https://www6.fgv.br/tic/identitymanager/unifiedlogon?id=eb1e7422-2d0f-45a2-ab02-f49bd82c8e8d&returnurl=https://sigaiwi.fgv.br/acad/hub/account/identitymanagerlogin?token=')

        # Logar no SIGA
        pagina.wait_for_selector('#user').fill('rodrigo.vasconcelos')
        pagina.wait_for_selector('#password').fill('Rdg#2024')
        pagina.keyboard.press('Enter')

        # Acessar o menu
        pagina.wait_for_selector('//span[text()=" Financeiro "]').click()
        pagina.wait_for_selector('//span[text()=" Plano Financeiro Presencial "]').click()

        # Buscar a matrícula do aluno
        frame_PlanFinPresencial = pagina.frame_locator('#frameSiga2')

        arquivo = openpyxl.load_workbook(caminho_arquivo)
        planilha = arquivo['Base']

        contador = 1

        for col in planilha.iter_cols(max_col=1, min_row=2):
            for cell in col:
                frame_PlanFinPresencial.locator('#SistemaContentPlaceHolder_BtnAbrePop').click()
                frame_PlanFinPresencial.locator('#ResultadoBuscaContentPlaceHolder_ucModalBuscarAluno_ctl00_BuscaMatriculaTextBox').fill(cell.value)
                frame_PlanFinPresencial.locator('#ResultadoBuscaContentPlaceHolder_ucModalBuscarAluno_ctl00_BuscaAlunoButton').click()
                frame_PlanFinPresencial.get_by_text(cell.value).click()

                # Acessar a tela de parcelas e gerar o relatório
                frame_PlanFinPresencial.locator('#__tab_SistemaContentPlaceHolder_tabContainerFinanceiro_Mensalidade').click()

                try:
                    with pagina.expect_download() as download_info:
                        frame_PlanFinPresencial.locator('#SistemaContentPlaceHolder_BtnGerarExcel').click()

                    download = download_info.value
                    celula_aluno = cell.offset(row=0, column=1).value
                    nome_aluno = str(celula_aluno).upper()
                    nome_arquivo = f'{nome_aluno} - {cell.value.replace("/", "_")}.xls'
                    caminho_download = os.path.join(r'F:\downloads', nome_arquivo)
                    download.save_as(caminho_download)

                except TimeoutError:
                    print("O botão de download não está disponível.")
                
                finally:
                    # Clicar no botão de voltar para gerar o próximo plano financeiro
                    frame_PlanFinPresencial.locator('#SistemaContentPlaceHolder_DadosAluno_BtnVoltar').click()
                    contador += 1
        
        navegador.close()

if __name__ == "__main__":
    caminho_arquivo = sys.argv[1]
    automacao(caminho_arquivo)