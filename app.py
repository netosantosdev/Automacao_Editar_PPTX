import pandas as pd
from pptx import Presentation
import os
import re
from pathlib import Path
import sys
import time

def sanitize_filename(filename):
    """Remove caracteres inválidos para nomes de arquivos"""
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def gerar_certificados(template_path, csv_path, output_folder):
    # Verificações iniciais
    print(f"\nVerificando arquivos...")
    print(f"Template existe? {os.path.exists(template_path)}")
    print(f"CSV existe? {os.path.exists(csv_path)}")
    
    # Cria certificados_gerados caso não exista
    Path(output_folder).mkdir(parents=True, exist_ok=True)
    
    # Ler CSV
    try:
        df = pd.read_csv(csv_path)
        print(f"\nTotal de registros encontrados: {len(df)}")
    except Exception as e:
        print(f"\nERRO ao ler CSV: {e}")
        return

    # Contadores para relatório
    sucessos = 0
    erros = 0

    # Processar cada registro
    for index, row in df.iterrows():
        try:
            nome = str(row['nome']).strip()
            numero = str(row['numero']).strip()
            
            print(f"\nProcessando: {nome} (Certificado {numero})")
            
            # Sanitizar valores
            nome_safe = sanitize_filename(nome)
            numero_safe = sanitize_filename(numero)
            
            # Carregar template
            prs = Presentation(template_path)
            
            # Substituir placeholders
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.text = run.text.replace('{NOME}', nome)
                                run.text = run.text.replace('{NUMERO}', numero)
            
            # Criar nome de arquivo seguro
            output_filename = f"Certificado_{numero_safe}_{nome_safe[:50]}.pptx"
            output_path = os.path.join(output_folder, output_filename)
            
            # Verificar caminho final
            print(f"Tentando salvar em: {output_path}")
            
            # Salvar
            prs.save(output_path)
            print(f"→ Certificado gerado com sucesso!")
            sucessos += 1
            
        except Exception as e:
            print(f"ERRO ao processar linha {index + 1}: {e}")
            erros += 1
    
    # Relatório final
    print("\n" + "="*50)
    print("RELATÓRIO DE EXECUÇÃO")
    print(f"Total de certificados processados: {len(df)}")
    print(f"Certificados gerados com sucesso: {sucessos}")
    print(f"Erros encontrados: {erros}")
    print("="*50)

def manter_prompt_aberto():
    """Mecanismo para manter o prompt aberto no final da execução"""
    if sys.platform.startswith('win'):
        # Para Windows
        print("\nPressione ENTER para fechar esta janela...")
        input()
    else:
        # Para outros sistemas (Linux/Mac)
        print("\nO terminal fechará automaticamente em 30 segundos...")
        time.sleep(30)

if __name__ == "__main__":
    # CONFIGURAÇÕES - AJUSTE ESTES CAMINHOS!
    TEMPLATE_PATH = "./base/certificado_template.pptx"
    CSV_PATH = "./base/dados_participantes.csv"
    OUTPUT_FOLDER = "./certificados_gerados"
    
    print("Iniciando processo de geração de certificados...")
    gerar_certificados(TEMPLATE_PATH, CSV_PATH, OUTPUT_FOLDER)
    
    # Manter o prompt aberto
    manter_prompt_aberto()