import csv
import os
import json
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def rename_arquivo(participantes):
    numero = 1
    for participante in participantes:    
        antes = r'ilovepdf_split\\Todos-' + str(numero) + '.pdf'
        depois = r'ilovepdf_split\\Brawl 2023 - ' + participante + '.pdf'
        os.rename(antes, depois)
        print("Renomeado o arquivo" + str(numero))
        numero = numero + 1

def carrega_dados(caminho):
    dados_socios = {}
    with open(caminho, "r", encoding="utf-8") as f:
        for x in json.load(f): dados_socios[x["nome"]] = x 
    return dados_socios

def get_dados_socios(participantes, dados_socios):
    for participante in participantes:
        if participante in dados_socios.keys():
           dados_socios[participante]["participante_brawl"] = True
        else:
            print("Não encontrei x participante {}".format(participante))
    return dados_socios

def confere_certificados(dados_socios):
    for socio in dados_socios:
        if "participante_brawl" not in socio.keys(): continue
        if not os.path.isfile(r'ilovepdf_split\\Brawl 2023 - {}.pdf'.format(socio["nome"])):
            print("Não encontrei o certificado dx {}".format(socio["nome"]))

with open("participantes.csv", "r", encoding="utf-8-sig") as f:
    participantes = [x["COMPETIDOR"] for x in csv.DictReader(f, delimiter=";")]

def constroi_mensagem(socio):
    msg = MIMEMultipart()
    msg['Subject'] = "ABRACM - Certificado Brasília Brawl 2023"
    corpo = "Prezad{} Associad{},\n\nObrigado por participar do Brasília Brawl 2023! \
Sua presença foi muito importante para um dia muito divertido de disputas, reviravoltas \
e conquistas. Esperamos você na próxima competição!\n\nEnviamos em anexo o certificado \
da sua participação!".format(socio["artigo_genero"], socio["artigo_genero"])
    msg.attach(MIMEText(corpo))
    msg["To"] = socio["email"]
    file_name = r'ilovepdf_split\\Brawl 2023 - {}.pdf'.format(socio["nome"])
    with open(file_name, "rb") as f:
        anexo = MIMEApplication(
                                f.read(),
                                Name=os.path.basename(file_name)
                            )
    anexo['Content-Disposition'] = 'attachment; filename="%s"' % os.path.basename(file_name)
    msg.attach(anexo)
    return msg.as_string()

def get_senhas():
    with open("senhas.txt", "r", encoding="utf-8") as f:
        return [x.replace("\n", "") for x in f]

def envia_emails(dados_socios, envios_manuais=[]):
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login("contato@abracm.org.br", get_senhas()[0])
    for socio in dados_socios:
        if envios_manuais != [] and socio["nr_inscricao"] not in envios_manuais: continue
        try:
            server.sendmail("contato@abracm.org.br", socio["email"], constroi_mensagem(socio))
            socio["enviado"] = True
        except:
            print("{} - não deu certo o envio".format(socio["nome"]))
            socio["enviado"] = False
    server.quit()
    
dados_socios = carrega_dados("dados_socios.json")
dados_socios = get_dados_socios(participantes, dados_socios)
dados_socios = [dados_socios[key] for key in dados_socios]
confere_certificados(dados_socios)
envia_emails(dados_socios)
