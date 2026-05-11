from pydoc import doc
import flet as ft
import fitz  # PyMuPDF
from PIL import Image
import io, base64, os
from docx2pdf import convert
import pythoncom
import sys
import win32com.client
import shutil
import datetime



base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
caminho_word = os.path.join(base_path, "meuarquivo.docx")

paginas_pdf = []
        # Pilha para armazenar páginas removidas
paginas_removidas = []


def main(page: ft.Page):
    
    page.window_maximized = True
    page.theme_mode = ft.ThemeMode.DARK
   

   
    thumbnails = [None] * 40
    thumbnails_row = ft.Row(wrap=True, spacing=10)
    current_index = {"value": None}


    NUM_CAIXAS = 40
    word_files = [None] * NUM_CAIXAS
    word_row = ft.Row(wrap=True, spacing=10)


    def resource_path(relative_path):
        # Se estiver rodando empacotado, os arquivos ficam em _MEIPASS
        if hasattr(sys, "_MEIPASS"):
            return os.path.join(sys._MEIPASS, relative_path)
        # Caso contrário, usa o caminho normal
        return os.path.join(os.path.abspath("."), relative_path)
    
    logo_path = resource_path("assets/logodocflow.png")
    logo_image = ft.Image(logo_path, width=450, height=450)


    # ---------- Views ----------


   
    def home_view():
        
        return ft.View(
            "/",
            [
                ft.Container(
                    content=ft.Column(
                        [
                            logo_image,
                            
                            ft.Row([ft.ElevatedButton(
                                        content=ft.Container(
                                            content=ft.Column(
                                                [
                                                    ft.Text("           ", size=3),
                                                    ft.Text("Mesclar PDF", size=25, weight="bold", color=ft.colors.WHITE),
                                                    ft.Text("Junte vários arquivos PDF em um único documento.", size=13.5, color=ft.colors.WHITE),
                                                ],
                                                alignment=ft.MainAxisAlignment.START,
                                                horizontal_alignment=ft.CrossAxisAlignment.START,
                                                spacing=10,
                                            ),
                                            expand=True,
                                            alignment=ft.alignment.top_left,
                                        ),
                                        on_click=lambda _: page.go("/mesclar"),
                                        bgcolor=ft.colors.RED_500,
                                        width=350,
                                        height=100,
                                        style=ft.ButtonStyle(
                                            shape=ft.RoundedRectangleBorder(radius=30),
                                            side=ft.BorderSide(3, ft.colors.RED),
                                            overlay_color=ft.colors.RED_100,
                                            shadow_color=ft.colors.RED_900,
                                            elevation=5,  # sombra mais escura
                                        ),
                                    ),
                                    ft.ElevatedButton(
                                        content=ft.Container(
                                            content=ft.Column(
                                                [
                                                    ft.Text("           ", size=3),
                                                    ft.Text("Converter WORD para PDF", size=25, weight="bold", color=ft.colors.WHITE),
                                                    ft.Text("Transforme WORD em outros formatos, como PDF.", size=13.5, color=ft.colors.WHITE),
                                                ],
                                                alignment=ft.MainAxisAlignment.START,
                                                horizontal_alignment=ft.CrossAxisAlignment.START,
                                                spacing=10,
                                            ),
                                            expand=True,
                                            alignment=ft.alignment.top_left,
                                        ),
                                        on_click=lambda _: page.go("/converter"),
                                        bgcolor=ft.colors.BLUE_500,
                                        width=350,
                                        height=100,
                                        style=ft.ButtonStyle(
                                            shape=ft.RoundedRectangleBorder(radius=30),
                                            side=ft.BorderSide(3, ft.colors.BLUE_500),
                                            overlay_color=ft.colors.BLUE_100,
                                            shadow_color=ft.colors.BLUE_900,
                                            elevation=5,
                                        ),
                                    ),
                                    ft.ElevatedButton(
                                        content=ft.Container(
                                            content=ft.Column(
                                                [
                                                    ft.Text("           ", size=3),
                                                    ft.Text("Separar PDF", size=18, weight="bold", color=ft.colors.WHITE),
                                                    ft.Text("Divida um PDF em páginas ou partes específicas.", size=14, color=ft.colors.WHITE),
                                                ],
                                                alignment=ft.MainAxisAlignment.START,
                                                horizontal_alignment=ft.CrossAxisAlignment.START,
                                                spacing=10,
                                            ),
                                            expand=True,
                                            alignment=ft.alignment.top_left,
                                        ),
                                        on_click=lambda _: page.go("/separar"),
                                        bgcolor=ft.colors.GREY_500,
                                        width=300,
                                        height=100,
                                        style=ft.ButtonStyle(
                                            shape=ft.RoundedRectangleBorder(radius=30),
                                            side=ft.BorderSide(3, ft.colors.GREY_500),
                                            overlay_color=ft.colors.GREY_100,
                                        ),
                                    ),
                                ],
                                alignment=ft.MainAxisAlignment.CENTER,
                                spacing=40,
                            ),
                        ],
                        alignment=ft.MainAxisAlignment.START,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=50,
                    scroll=ft.ScrollMode.AUTO
                ),
                expand=True,
                alignment=ft.alignment.top_center,
                gradient=ft.LinearGradient(
                    begin=ft.alignment.top_right,
                    end=ft.alignment.bottom_right,
                    colors=[ft.colors.BLUE_GREY, ft.colors.WHITE],
                ),
            )
        ],
    )
    # Fundo com degradê
  
    Barra_nome_arquivo = ft.TextField(
                            hint_text="Nome do arquivo final (sem extensão)",
                            prefix_icon=ft.Icon(
                                name=ft.icons.INSERT_DRIVE_FILE,
                                color=ft.colors.BLACK,   # cor do ícone
                                size=30                 # opcional: tamanho do ícone
                            ),
                            hint_style=ft.TextStyle(color=ft.colors.GREY_500),
                            bgcolor=ft.colors.WHITE,
                            color=ft.colors.BLACK,
                            border_radius=10,
                            width=450,
                            border_color=ft.colors.WHITE,
                            focused_border_color=ft.colors.BLACK,
                            height=40,
                            text_size=15,
                            text_align=ft.TextAlign.LEFT,
                            content_padding=10
                        )

    def mesclar_view():
        return ft.View(
            "/mesclar",
            [
                ft.Row(
                    [
                        ft.ElevatedButton(
                            "Voltar",
                            icon=ft.icons.ARROW_BACK,   # ícone de retorno
                            bgcolor=ft.colors.RED,      # cor vermelha
                            color=ft.colors.WHITE,
                            on_click=lambda _: voltar_home()
                        ),
                        ft.Container(
                            content=Barra_nome_arquivo,
                            width=500
                        ),
                        ft.ElevatedButton(
                            "Exportar PDF na ordem",
                            icon=ft.icons.FILE_DOWNLOAD,     # ícone de salvar/baixar
                            bgcolor=ft.colors.GREEN,    # cor verde
                            color=ft.colors.WHITE,
                            on_click=exportar_pdfs
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                ft.Container(
                    content=ft.Column(
                        [thumbnails_row],
                        scroll=ft.ScrollMode.AUTO,   # scroll automático
                        expand=True
                    ),
                    expand=True
                )
            ],
        )

    def voltar_home():
        # limpa ao voltar
        thumbnails[:] = [None] * 40
        render_thumbnails()
        page.go("/")

   
    def converter_view():
        render_word_boxes()
        return ft.View(
            "/converter",
            [
                ft.Row(
                    [
                        ft.ElevatedButton(
                            "Voltar",
                            icon=ft.icons.ARROW_BACK,
                            bgcolor=ft.colors.RED,
                            color=ft.colors.WHITE,
                            on_click=lambda _: page.go("/")
                        ),
                        ft.ElevatedButton(
                            "Converter Word → PDF",
                            icon=ft.icons.PICTURE_AS_PDF,
                            bgcolor=ft.colors.BLUE_700,
                            color=ft.colors.WHITE,
                            on_click=converter_word_para_pdf
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                ft.Container(
                    content=ft.Column([word_row], scroll=ft.ScrollMode.AUTO, expand=True),
                    expand=True
                )
            ],
        )

    def render_word_boxes():
        word_row.controls.clear()
        for idx, item in enumerate(word_files):
            if item:
                nome, path = item
                conteudo = ft.Column(
                    [
                        ft.Row(
                            [
                                ft.IconButton(
                                    icon=ft.icons.CLOSE,
                                    icon_color=ft.colors.RED,
                                    tooltip="Excluir",
                                    on_click=lambda _, i=idx: excluir_word(i)
                                )
                            ],
                            alignment=ft.MainAxisAlignment.END
                        ),
                        ft.Text(nome, size=12, weight=ft.FontWeight.BOLD),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER
                )
            else:
                conteudo = ft.Column(
                    [
                        ft.Text(f"Posição {idx+1}", size=12, weight=ft.FontWeight.BOLD),
                        ft.Text("Nenhum Word"),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER
                )

            botao_upload = ft.ElevatedButton(
                "Upload Word",
                on_click=lambda _, i=idx: escolher_word(i)
            )

            caixa = ft.Container(
                content=ft.Column([conteudo, botao_upload], spacing=10),
                width=160,
                height=120,
                border=ft.border.all(2, ft.colors.BLUE_700),
                border_radius=10,
                padding=10,
                margin=5,
                alignment=ft.alignment.center
            )
            word_row.controls.append(caixa)

        page.update()

    def excluir_word(index):
        if 0 <= index < len(word_files):
            word_files[index] = None
            render_word_boxes()

    def escolher_word(index):
        current_index["value"] = index
        file_picker_word.pick_files(allowed_extensions=["docx"], allow_multiple=True)

    def on_result_word(e: ft.FilePickerResultEvent):
        idx = current_index["value"]
        if e.files and idx is not None:
            for f in e.files:
                for i in range(len(word_files)):
                    if word_files[i] is None:
                        word_files[i] = (f.name, f.path)
                        break
            render_word_boxes()
    def converter_word_para_pdf(e=None, arquivos=None):
        """
        Se chamada pelo botão do Flet, ignora 'e' e usa word_files.
        Se chamada pelo Prompt de Comando, usa 'arquivos' passados por sys.argv.
        """
        if arquivos is None:  # caso interface Flet
            arquivos = [item[1] for item in word_files if item]

        if not arquivos:
            print("Nenhum Word para converter.")
            page.snack_bar = ft.SnackBar(ft.Text("Nenhum arquivo Word selecionado."))
            page.snack_bar.open = True
            page.update()
            return

        try:
            pythoncom.CoInitialize()

            downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(downloads_dir, exist_ok=True)

            print(f"Iniciando conversão de {len(arquivos)} arquivo(s)...")
            for caminho in arquivos:
                print(f"Convertendo: {caminho}")
                word = win32com.client.Dispatch("Word.Application")
                doc = word.Documents.Open(caminho)

                nome_pdf = os.path.splitext(os.path.basename(caminho))[0] + ".pdf"
                caminho_pdf = os.path.join(downloads_dir, nome_pdf)

                doc.SaveAs(caminho_pdf, FileFormat=17)
                doc.Close()
                word.Quit()

                print(f"✔ Arquivo convertido e salvo em {downloads_dir}")

            # limpa lista e atualiza interface (arquivos somem da tela)
            word_files[:] = [None] * NUM_CAIXAS
            render_word_boxes()

            print("✅ Conversão concluída com sucesso.")

            # só depois de limpar a tela, mostra a mensagem
            page.snack_bar = ft.SnackBar(
                ft.Text("Conversão concluída com sucesso! PDF salvo na pasta Downloads.")
            )
            page.snack_bar.open = True
            page.update()

        except Exception as err:
            print(f"❌ Erro ao converter: {err}")
            with open("erro_log.txt", "a", encoding="utf-8") as log:
                log.write(str(err) + "\n")

            page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao converter: {err}"))
            page.snack_bar.open = True
            page.update()
        finally:
            pythoncom.CoUninitialize()


    # Inicialização do FilePicker específico para Word
    file_picker_word = ft.FilePicker(on_result=on_result_word)
    page.overlay.append(file_picker_word)

    # Execução via Prompt de Comando
    if __name__ == "__main__":
        if len(sys.argv) > 1:
            arquivos = sys.argv[1:]
            converter_word_para_pdf(arquivos=arquivos)
        else:
            print("Uso: TOOLS.exe arquivo1.docx arquivo2.docx ...")


    
    def organizar_pdf(idx):
        global paginas_pdf, paginas_removidas
        nome, _, path = thumbnails[idx]

        doc = fitz.open(path)
        paginas_pdf = list(range(len(doc)))
        paginas_removidas = []

        def render_paginas():
            controles = []
            for pos, page_num in enumerate(paginas_pdf):
                pix = doc[page_num].get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
                img_b64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")

                controles.append(
                    ft.Row(
                        [
                            ft.Image(src_base64=img_b64, width=100, height=140, fit=ft.ImageFit.CONTAIN),
                            ft.Column(
                                [
                                    ft.IconButton(
                                        icon=ft.icons.ARROW_UPWARD,
                                        tooltip="Mover para cima",
                                        on_click=lambda e, p=pos: mover_pagina(p, -1)
                                    ),
                                    ft.IconButton(
                                        icon=ft.icons.ARROW_DOWNWARD,
                                        tooltip="Mover para baixo",
                                        on_click=lambda e, p=pos: mover_pagina(p, 1)
                                    ),
                                    ft.IconButton(
                                        icon=ft.icons.DELETE,
                                        icon_color=ft.colors.RED,
                                        tooltip="Excluir página",
                                        on_click=lambda e, p=pos: excluir_pagina(p)
                                    ),
                                ],
                                spacing=2
                            )
                        ],
                        spacing=10,
                        alignment=ft.MainAxisAlignment.START
                    )
                )
            return controles

        def mover_pagina(pos, delta):
            new_pos = pos + delta
            if 0 <= new_pos < len(paginas_pdf):
                paginas_pdf[pos], paginas_pdf[new_pos] = paginas_pdf[new_pos], paginas_pdf[pos]
                atualizar()

        def excluir_pagina(pos):
            if 0 <= pos < len(paginas_pdf):
                # Primeiro guarda a página que será removida
                pagina_removida = paginas_pdf[pos]
                paginas_removidas.append((pos, pagina_removida))

                # Agora remove da lista
                paginas_pdf.pop(pos)

                # Recria documento atualizado
                novo_doc = fitz.open()
                for page_num in paginas_pdf:
                    novo_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

                # Caminho da pasta provisória
                downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
                pasta = os.path.join(downloads_dir, "pdfs_temp")
                os.makedirs(pasta, exist_ok=True)

                # Define caminho do novo arquivo dentro da pasta
                nome_arquivo = os.path.basename(path).replace(".pdf", "_organizado.pdf")
                novo_path = os.path.join(pasta, nome_arquivo)

                # Salva documento atualizado
                novo_doc.save(novo_path)
                novo_doc.close()

                # Atualiza thumbnail na posição com o arquivo modificado
                thumb_pix = fitz.open(novo_path)[0].get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
                thumb_b64 = base64.b64encode(thumb_pix.tobytes("png")).decode("utf-8")
                thumbnails[idx] = (nome, thumb_b64, novo_path)

                # Redesenha interface
                atualizar()
                render_thumbnails()
                page.snack_bar = ft.SnackBar(ft.Text(f"Página excluída. PDF atualizado em {novo_path}"))
                page.snack_bar.open = True
                page.update()

        def desfazer():
            if paginas_removidas:
                pos, pagina = paginas_removidas.pop()
                paginas_pdf.insert(pos, pagina)
                atualizar()

        def salvar_pdf():
            novo_doc = fitz.open()
            for page_num in paginas_pdf:
                novo_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)

            # Cria pasta provisória dentro da pasta Downloads
            downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            pasta = os.path.join(downloads_dir, "pdfs_temp")
            os.makedirs(pasta, exist_ok=True)

            # Define caminho do novo arquivo dentro da pasta Downloads/pdfs_temp
            nome_arquivo = os.path.basename(path).replace(".pdf", "_organizado.pdf")
            novo_path = os.path.join(pasta, nome_arquivo)

            # Salva o PDF reorganizado na pasta provisória
            novo_doc.save(novo_path)
            novo_doc.close()

            # Atualiza thumbnail na posição com o arquivo modificado
            thumb_pix = fitz.open(novo_path)[0].get_pixmap(matrix=fitz.Matrix(0.2, 0.2))
            thumb_b64 = base64.b64encode(thumb_pix.tobytes("png")).decode("utf-8")
            thumbnails[idx] = (nome, thumb_b64, novo_path)

            # Fecha diálogo e redesenha thumbnails
            page.dialog.open = False
            render_thumbnails()
            page.snack_bar = ft.SnackBar(ft.Text(f"PDF atualizado em {novo_path}"))
            page.snack_bar.open = True
            page.update()

        def atualizar():
            dialog.content.controls = render_paginas()
            page.update()

        def fechar_dialog():
            page.dialog.open = False
            page.update()

        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                [
                    ft.Text("Organizar PDF"),
                    ft.IconButton(
                        icon=ft.icons.CLOSE,
                        tooltip="Fechar",
                        on_click=lambda e: fechar_dialog()
                    )
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN
            ),
            content=ft.Column(
                controls=render_paginas(),
                scroll=ft.ScrollMode.AUTO,
                spacing=10
            ),
            actions=[
                ft.TextButton("Desfazer", on_click=lambda e: desfazer()),
                ft.TextButton("Salvar", on_click=lambda e: salvar_pdf())
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.dialog = dialog
        dialog.open = True
        page.update()


        
    def separar_view():
        return ft.View(
            "/separar",
            [
                ft.ElevatedButton("Voltar", on_click=lambda _: page.go("/")),
                ft.Text("Tela de separação (em construção)", size=18),
            ],
        )
    # ---------- Funções ----------


    def mover_caixa(idx, direcao):
        novo_idx = idx + direcao

        # Verifica se o novo índice é válido
        if 0 <= novo_idx < len(thumbnails):
            # Troca os elementos de posição
            thumbnails[idx], thumbnails[novo_idx] = thumbnails[novo_idx], thumbnails[idx]

            # Atualiza a interface
            render_thumbnails()

    # Use Wrap para permitir quebra automática de linha
    # Use ResponsiveRow para permitir quebra automática
    thumbnails_row = ft.ResponsiveRow(
        controls=[],
        alignment=ft.MainAxisAlignment.START,
     
        spacing=5,
        
    )

    def render_thumbnails():
        thumbnails_row.controls.clear()
        for idx, item in enumerate(thumbnails):
            if item:
                nome, img_b64, path = item
                conteudo = ft.Column(
                    [
                        ft.Row(
                            [
                                ft.IconButton(
                                    icon=ft.icons.CLOSE,
                                    icon_color=ft.colors.RED,
                                    tooltip="Excluir",
                                    on_click=lambda _, i=idx: excluir_caixa(i)
                                ),
                                ft.ElevatedButton(
                                    content=ft.Text("Organizar PDF", size=10),
                                    on_click=lambda _, i=idx: organizar_pdf(i)
                                )
                            ],
                            alignment=ft.MainAxisAlignment.END,
                            spacing=4
                        ),
                        ft.Text(nome, size=12, weight=ft.FontWeight.BOLD),
                        ft.Image(src_base64=img_b64, width=120, height=160, fit=ft.ImageFit.CONTAIN),
                    ],
                    spacing=4,
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER
                )
            else:
                conteudo = ft.Column(
                    [
                        ft.Text(f"Posição {idx+1}", size=12, weight=ft.FontWeight.BOLD),
                        ft.Text("Nenhum PDF"),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER
                )

            botao_upload = ft.ElevatedButton(
                "Upload PDF",
                on_click=lambda _, i=idx: escolher_pdf(i)
            )

            # Caixa acoplada (apenas PDF + botão)
            caixa = ft.Container(
                content=ft.Column(
                    [conteudo, botao_upload],
                    alignment=ft.MainAxisAlignment.CENTER,
                    horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                    spacing=4
                ),
                width=160,
                height=240,
                border=ft.border.all(2, ft.colors.RED),
                border_radius=8,
                padding=6,
                margin=6,
                alignment=ft.alignment.center
            )

            # Setas fora da caixa
            organizacao = ft.Row(
                [
                    ft.IconButton(
                    icon=ft.icons.ARROW_BACK,
                    icon_color=ft.colors.WHITE,
                    bgcolor=ft.colors.BLUE_700,
                    hover_color=ft.colors.BLUE_900,      # cor quando passa o mouse
                    highlight_color=ft.colors.BLACK,
                    tooltip="Mover para esquerda",
                        
                 
                        on_click=lambda _, i=idx: mover_caixa(i, -1)
                    ),
                    ft.IconButton(
                    icon=ft.icons.ARROW_FORWARD,
                    icon_color=ft.colors.WHITE,
                    bgcolor=ft.colors.BLUE_700,
                    hover_color=ft.colors.BLUE_900,      # cor quando passa o mouse
                    highlight_color=ft.colors.BLACK,    # cor quando está "pressionado"
                    tooltip="Mover para direita",
                    on_click=lambda _, i=idx: mover_caixa(i, 1)
                )

                ],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=4
            )

            # Junta caixa + setas em um Row (setas ficam fora da caixa)
            bloco = ft.Row(
                [caixa, organizacao],
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=8
            )

            # Responsividade
            thumbnails_row.controls.append(
                ft.ResponsiveRow(
                    controls=[bloco],
                    col={"xs": 12, "sm": 6, "md": 4, "lg": 3}
                )
            )

        page.update()








    def excluir_caixa(index):
        if 0 <= index < len(thumbnails):
            thumbnails[index] = None
            render_thumbnails()

    def escolher_pdf(index):
        current_index["value"] = index
        # permitir múltiplos arquivos
        file_picker.pick_files(allowed_extensions=["pdf"], allow_multiple=True)

    def on_result(e: ft.FilePickerResultEvent):
        idx = current_index["value"]
        if e.files and idx is not None:
            for f in e.files:  # percorre todos os PDFs escolhidos
                try:
                    doc = fitz.open(f.path)
                    if doc.page_count > 0:
                        pix = doc[0].get_pixmap(matrix=fitz.Matrix(2, 2))
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        img.thumbnail((200, 200))

                        buf = io.BytesIO()
                        img.save(buf, format="PNG")
                        img_b64 = base64.b64encode(buf.getvalue()).decode("utf-8")

                        # encontra a próxima posição livre
                        for i in range(len(thumbnails)):
                            if thumbnails[i] is None:
                                thumbnails[i] = (f.name, img_b64, f.path)
                                break
                    render_thumbnails()
                except Exception as err:
                    page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao abrir {f.name}: {err}"))
                    page.snack_bar.open = True
                    page.update()

##########################################################################################################################################################


    # FilePicker apenas para escolher a pasta destino
    file_picker_pasta = ft.FilePicker(on_result=lambda e: mover_arquivo(e))
    page.overlay.append(file_picker_pasta)

    ultimo_arquivo_exportado = None

    def exportar_pdfs(e):
        global ultimo_arquivo_exportado

        arquivos = [item[2] for item in thumbnails if item]
        if not arquivos:
            page.snack_bar = ft.SnackBar(ft.Text("Nenhum PDF para exportar"))
            page.snack_bar.open = True
            page.update()
            return

        try:
            novo_pdf = fitz.open()
            for caminho in arquivos:
                doc = fitz.open(caminho)
                novo_pdf.insert_pdf(doc)
                doc.close()

            downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")

            nome_digitado = Barra_nome_arquivo.value.strip()
            if nome_digitado:
                nome_arquivo = f"{nome_digitado}.pdf"
            else:
                timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"Arquivo_final_mesclado_{timestamp}.pdf"

            output_path = os.path.join(downloads_dir, nome_arquivo)

            novo_pdf.save(output_path)
            novo_pdf.close()

            # Guarda o caminho do arquivo exportado
            ultimo_arquivo_exportado = output_path

            # Limpa o campo de nome do arquivo
            Barra_nome_arquivo.value = ""
            page.update()

            thumbnails[:] = [None] * 40
            render_thumbnails()

            page.snack_bar = ft.SnackBar(ft.Text(f"Exportado para {output_path}. Agora escolha a pasta destino."))
            page.snack_bar.open = True
            page.update()

            # Abre automaticamente o seletor de pasta
            file_picker_pasta.get_directory_path()

        except Exception as err:
            page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao exportar: {err}"))
            page.snack_bar.open = True
            page.update()


    def mover_arquivo(e):
        global ultimo_arquivo_exportado
        if ultimo_arquivo_exportado and e.control.result and e.control.result.path:
            pasta_destino = e.control.result.path
            destino_final = os.path.join(pasta_destino, os.path.basename(ultimo_arquivo_exportado))
            shutil.move(ultimo_arquivo_exportado, destino_final)

            page.snack_bar = ft.SnackBar(ft.Text(f"Arquivo movido para {destino_final}"))
            page.snack_bar.open = True
            page.update()



    def mover_arquivo(e):
        global ultimo_arquivo_exportado

        if not ultimo_arquivo_exportado:
            page.snack_bar = ft.SnackBar(ft.Text("Nenhum arquivo exportado para mover"))
            page.snack_bar.open = True
            page.update()
            return

        try:
            pasta_destino = e.control.result.path
            if pasta_destino:
                destino_final = os.path.join(pasta_destino, os.path.basename(ultimo_arquivo_exportado))
                shutil.move(ultimo_arquivo_exportado, destino_final)

                page.snack_bar = ft.SnackBar(ft.Text(f"Arquivo movido para {destino_final}"))
                page.snack_bar.open = True
                page.update()
        except Exception as err:
            page.snack_bar = ft.SnackBar(ft.Text(f"Erro ao mover: {err}"))
            page.snack_bar.open = True
            page.update()

##########################################################################################################################################################

    # ---------- FilePicker global ----------
    file_picker = ft.FilePicker(on_result=on_result)
    page.overlay.append(file_picker)

    file_picker_word = ft.FilePicker(on_result=on_result_word)
    page.overlay.append(file_picker_word)

    # Inicializa thumbnails
    render_thumbnails()

    # ---------- Roteamento ----------
    def route_change(e):
        if page.route == "/":
            page.views.clear()
            page.views.append(home_view())
        elif page.route == "/mesclar":
            page.views.clear()
            page.views.append(mesclar_view())
        elif page.route == "/converter":
            page.views.clear()
            page.views.append(converter_view())
        elif page.route == "/separar":
            page.views.clear()
            page.views.append(separar_view())
        page.update()

    page.on_route_change = route_change
    page.go("/")

ft.app(target=main)

#ft.app(target=main, view=ft.AppView.WEB_BROWSER)

#flet pack DocFlow.py --icon assets/iconeflow.png --add-data "assets;assets" --name DocFlow


