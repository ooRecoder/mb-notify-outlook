import subprocess
import sys
import os
import shutil
import argparse

def limpar_diretorios():
    """
    Remove as pastas geradas em compilações anteriores.
    """
    pastas = ("build", "dist", "__pycache__")
    for pasta in pastas:
        if os.path.exists(pasta):
            print(f"🧹 Removendo pasta antiga: {pasta}")
            shutil.rmtree(pasta)
    print("✅ Limpeza concluída!\n")

def compilar():
    """
    Executa o PyInstaller para gerar o executável único.
    """
    comando = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--name", "CriarLembretes (BETA)",
        "--hidden-import=pandas",
        "--hidden-import=numpy",
        "--hidden-import=openpyxl",
        "--hidden-import=win32com",
        "--hidden-import=win32com.client",
        "--hidden-import=win32com.server",
        "--hidden-import=win32com.shell",
        "--hidden-import=win32timezone",
        "src/main.py"
    ]

    print("🛠️ Iniciando compilação com PyInstaller...\n")
    subprocess.run(comando, check=True)
    print("\n✅ Compilação concluída com sucesso!")
    print("📦 Executável disponível em: dist/LimparLembretes.exe\n")

def main():
    parser = argparse.ArgumentParser(
        description="Script para compilar o projeto LimparLembretes com PyInstaller."
    )
    parser.add_argument(
        "--sem-limpar",
        action="store_true",
        help="Pula a limpeza dos diretórios de build antes da compilação."
    )
    parser.add_argument(
        "--somente-limpar",
        action="store_true",
        help="Apenas limpa os diretórios, sem compilar."
    )
    args = parser.parse_args()

    raiz = os.path.dirname(os.path.abspath(__file__))
    os.chdir(raiz)

    # Modo somente limpeza
    if args.somente_limpar:
        limpar_diretorios()
        return

    # Limpeza antes da compilação (a menos que o usuário peça para pular)
    if not args.sem_limpar:
        limpar_diretorios()
    else:
        print("⚙️ Pulando limpeza de diretórios.\n")

    compilar()

if __name__ == "__main__":
    main()
