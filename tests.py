import os, sys
import flet as ft

def resource_path(relative_path):
    # Se estiver rodando empacotado, os arquivos ficam em _MEIPASS
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    # Caso contrário, usa o caminho normal
    return os.path.join(os.path.abspath("."), relative_path)

def main(page: ft.Page):
    logo_path = resource_path("assets/logodocflow.png")
    logo_image = ft.Image(logo_path, width=300, height=300)

    page.add(
        ft.Column(
            [
                ft.Container(content=logo_image, alignment=ft.alignment.top_center),
                ft.Text("Bem-vindo ao DocFlow", size=22, weight="bold"),
            ],
            alignment=ft.MainAxisAlignment.START,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
            expand=True,
        )
    )

ft.app(target=main)

#flet pack DocFlow.py --icon assets/logodocflow.ico --add-data "assets;assets" --name DocFlow
