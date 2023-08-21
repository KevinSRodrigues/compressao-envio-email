import os
from zipfile import ZipFile, ZIP_DEFLATED
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askopenfilename
import win32com.client as win32


def compactar_tudo(diretorio, ignore_zips=True):
    nomesarquivo = os.listdir(diretorio)
    if ignore_zips:
        nomesarquivo = [fn for fn in nomesarquivo if not fn.endswith(".zip")]

    for nome in nomesarquivo:
        fullpath = os.path.join(diretorio, nome)
        if os.path.isdir(fullpath):
            nomezip = os.path.join(diretorio, nome + ".zip")
            arquivozip = ZipFile(nomezip, "a", compression=ZIP_DEFLATED)
            for raiz, dirs, arquivos in os.walk(fullpath):
                for arq in arquivos:
                    relativo = os.path.relpath(raiz, diretorio)
                    arquivozip.write(os.path.join(raiz, arq),
                                     os.path.join(relativo, arq))

        else:
            semextensao = nome.split(".")[0]
            nomezip = os.path.join(diretorio, semextensao + ".zip")
            arquivozip = ZipFile(nomezip, "w", compression=ZIP_DEFLATED)
            arquivozip.write(fullpath, nome)
            arquivozip.close()
    return len(nomesarquivo)


if __name__ == "__main__":
    initialdir = "c:\\Users\\Dell\\Desktop"
    diretorio = askdirectory(
        title="Selecione a pasta para compactar", initialdir=initialdir)
    print(f"Compactando arquivos na pasta {diretorio}")
    n = compactar_tudo(diretorio)
    print(f"{n} arquivos compactados com sucesso")


def enviar_email(outlook):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = "Destinatário"
    mail.Subject = "Teste"

    mail.Attachments.Add(attachment_path)

    pathToIMage = askopenfilename(title="Selecione sua assinatura")
    attachment = mail.Attachments.Add(pathToIMage)

    attachment.PropertyAccessor.SetProperty(
        "http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")

    mail.HTMLBody = '''

    <p>Teste</p>

    <p><figure><img src="cid:MyId1"></figure></p>

    '''
    mail.display()
    mail.Send()


if __name__ == "__main__":
    attachment_path = askopenfilename(
        title="Selecione o arquivo", initialdir=diretorio)
    b = enviar_email(attachment_path)
    print(f"Arquivo {attachment_path} anexado e e-mail enviado com sucesso!")
