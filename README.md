# chatbot-para-service-desk
um chatbot para atendimento que envia e-mail e cria planilhas de relatório do atendimento

# Criado e desenvolvido por: Marcelo Pugan


from datetime import datetime
import os
import smtplib
import openpyxl
import ast 
from email.message import EmailMessage

data = datetime.now()

EMAIL_ADDRESS = 'puganchatbot@gmail.com'
EMAIL_PASSWORD = '************'

nome = input('qual o seu nome? ')
matricula = input('qual a sua matrícula? ')
email_colaborador = input('Pode me informar um e-mail para contato? ')
print('\nOlá,',nome ,matricula, ', eu sou o chatbot do service desk, em que posso lhe ajudar? \n\n')
while True:    
    print('Você teria problema com: ')
    resposta = input('[1]hardware ou \n[2]software? \n\n')
    if(resposta == '1'): 
        print('O seu problema com hardware seria com:\n ')
        resposta2 = input('[1]periféricos, no \n[2]gabinete ou \n[3]celular? \n\n')
        if(resposta2 == '1'):
            print('O problema com um periférico é de: \n')
            resposta4 = input('[1]entrada ou de \n[2]saída? \n\n')
            if(resposta4 == '1'):
                print('O periférico em questão é um \n')
                periferico1 = input('[1]mouse, \n[2]teclado, \n[3]microfone ou \n[4]webcam? \n\n')
                if(periferico1 == '1'):
                    print('Será aberto um chamado para a substituição do seu mouse. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([ data, nome, matricula, 'troca de mouse', 'muito bom'])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de mouse', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de mouse', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de mouse', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de mouse', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(periferico1 == '2'):
                    print('Será aberto um chamado para a substituição do seu teclado. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de teclado', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de teclado', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de teclado', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de teclado', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de teclado', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break    
                if(periferico1 == '3'):
                    print('Será aberto um chamado para a substituição do seu microfone. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de microfone', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de microfone', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de microfone', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de microfone', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de microfone', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break     
                if(periferico1 == '4'):
                    print('Será aberto um chamado para a substituição do seu webcan. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de webcan', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de webcan', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de webcan', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de webcan', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de webcan', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break    
            elif(resposta4 == '2'):
                print('O periférico em questão é um \n')
                periferico2 = input('[1]monitor, \n[2]caixa de som, \n[3]fone de ouvido ou \n[4]impressora? \n\n')
                if(periferico2 == '1'):
                    print('Será aberto um chamado para a substituição do seu monitor. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de monitor', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de monitor', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de monitor', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de monitor', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de monitor', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break 
                if(periferico2 == '2'):
                    print('Será aberto um chamado para a substituição da sua caixa de som. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de caixa de som', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de caixa de som', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de caixa de som', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de caixa de som', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de caixa de som', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break     
                if(periferico2 == '3'):
                    print('Será aberto um chamado para a substituição do seu fone de ouvido. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de caixa de som', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de fone de ouvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de fone de ouvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de fone de ouvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de fone de ouvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break    
                if(periferico2 == '4'):
                    print('Será aberto um chamado para a manutenção da impressora. \n')
                    print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de impressora', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de impressora', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de impressora', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de impressora', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'troca de impressora', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break     
                    # Criado e desenvolvido por: Marcelo Pugan
        elif(resposta2 == '2'):
            print('O gabinete em questão: \n')
            resposta5 = input('[1]não está ligando, ou está \n[2]apresentando algum efeito estranho? \n\n')
            if(resposta5 == '1'):
                print('verifique se o computador está conectado na tomada, e tente noventente. \n')
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    conclusao_ruim=input('[1]Prefere recomeçar? ou\n[2]Finalizar o atendimento? \n\n')
                    if(conclusao_ruim == '1'):
                        print('Recomeçando... \n')
                    if(conclusao_ruim == '2'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break    
                    # Criado e desenvolvido por: Marcelo Pugan
            elif(resposta5 == '2'):
                print('Verifique se o ambiente em que fica localizado o computador acumula muita poeira, e faça uma limpeza em seu computador. \n')        
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    conclusao_ruim=input('[1]Prefere recomeçar? ou\n[2]Finalizar o atendimento? \n\n')
                    if(conclusao_ruim == '1'):
                        print('Recomeçando... \n')
                    if(conclusao_ruim == '2'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break    
        elif(resposta2 == '3'):
            print('O seu celular é um aparelho: \n')
            aparelho = input('[1]android, \n[2]IOS ou \n[3]analogico? \n\n')
            if(aparelho =='1'):
                print('desligue e ligue o celular. \n')
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    conclusao_ruim=input('[1]Prefere recomeçar? ou\n[2]Finalizar o atendimento? \n\n')
                    if(conclusao_ruim == '1'):
                        print('Recomeçando... \n')
                    if(conclusao_ruim == '2'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
            elif(aparelho=='2'):
                print('Se você tem dinheiro para ter um produto Apple, você tem dinheiro para ter um novo. \n')
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    # Criado e desenvolvido por Marcelo Pugan
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
            elif(aparelho=='3'):
                print('Se você ainda tem um celular analogico, esse chatbot, não pode lhe ajudar. \n')
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')        
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
    elif (resposta == '2'): 
        print('O seu problema com software, seria: \n')    
        resposta3 = input('[1]sistema operacional ou algum \n[2]programa? \n\n')
        if(resposta3 == '1'):
            print('Qual sistiema opercional seria: \n')
            resposta6 = input('[1]windows, \n[2]linux ou \n[3]IOS? \n\n')
            # Criado e desenvolvido por: Marcelo Pugan
            if(resposta6 =='1'):
                print('desligue e ligue o computador. \n')
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    conclusao_ruim=input('Prefere [2]Finalizar o atendimento? ou \n[1]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
            elif(resposta6 =='2'):
                print('Se você usa sistema operacional Linux, você provavelmente você sabe mais que esse chatbot. \n')
                print('Essa orientação lhe ajudou?\n')
                # Criado e desenvolvido por: Marcelo Pugan
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        # Criado e desenvolvido por: Marcelo Pugan
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
            elif(resposta6 =='3'):    
                print('Se você tem dinheiro para ter um produto Apple, você tem dinheiro para ter um novo. \n')
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    # Criado e desenvolvido por: Marcelo Pugan
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
        elif(resposta3 == '2'):
            print('O programa em questão: é uma requisição para \n')
            resposta7 = input('[1]instalação ou um incidente com \n[2]mal funcionamento? \n\n')
            if(resposta7 == '1'):
                print('Sobre o programa em questão. ')
                programa_instalacao = input('Escreva por favor o nome e a versão: \n\n')
                print('abriremos um chamado para realizar a instalação do ',programa_instalacao)
                print('\nNesse caso, vamos concluir o cadastro dos seus dados: \n')
                inventario = input('Qual o inventário do seu computador: \n')
                localicade = input('A a localização do seu setor? \n')
                telefone = input('Informe, por favor, um telefone de contato: \n')
                print('Será aberto um chamado com os seguintes dados: \n',nome,'\n',matricula,'\n',inventario,'\n',localicade,'\n',telefone,'\n',programa_instalacao )
                print('Essa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'puganchatbot.com'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    # Criado e desenvolvido por: Marcelo Pugan
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
            elif(resposta7 == '2'):    
                print('Sobre o programa em questão. ')
                programa_manutecao = input('Escreva por favor o nome e a versão: \n\n')
                print('abriremos um chamado para a manutençaõ do ',programa_manutecao)
                print('\nNesse caso, vamos concluir o cadastro dos seus dados: \n')
                inventario = input('Qual o inventário do seu computador: \n')
                localicade = input('A a localização do seu setor? \n')
                telefone = input('Informe, por favor, um telefone de contato: \n')
                print('Será aberto um chamado com os seguintes dados: ',data,'\n',nome,'\n',matricula,'\n',inventario,'\n',localicade,'\n',telefone,'\n',programa_manutecao )
                print('\nEssa orientação lhe ajudou?\n')
                conclusao = input('[s]im ou [n]ão \n')
                if(conclusao=='s'):
                    print('pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    # Criado e desenvolvido por: Marcelo Pugan
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
    else:
        print('Não entendi')
        nao_entendi = input('gostaria de \n[1]recomeçar ou \n[2]abrir um chamado? \n')
        if(nao_entendi == '1'):
            print('Recomeçando...')
        if(nao_entendi == '2'):
            print('\nNesse caso, vamos concluir o cadastro dos seus dados: \n')
            inventario = input('Qual o inventário do seu computador: \n')
            localicade = input('A a localização do seu setor? \n')
            telefone = input('Informe, por favor, um telefone de contato: \n')
            print('Será aberto um chamado com os seguintes dados: \n',data,'\n',nome,'\n',matricula,'\n',inventario,'\n',localicade,'\n',telefone )
            print('\nEssa orientação lhe ajudou?\n')
            conclusao = input('[s]im ou [n]ão \n')
            if(conclusao =='s'):
                print('pode avaliar como foi meu atendimento? \n')
                feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                if(feedback=='1'):    
                    relatorio = openpyxl.Workbook()
                    print(relatorio.sheetnames)
                    relatorio.create_sheet('relatorio do atendimento')
                    relatorio_page = relatorio['relatorio do atendimento']
                    relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito bom',])
                    relatorio.save('Relatorio de atendimento.xlsx')
                    msg = EmailMessage()
                    msg['Suject'] = 'Feedback do atendimento do chatbot'
                    msg['from'] = 'pugan chatbot'
                    msg['To'] = email_colaborador 
                    msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                        smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                        smtp.send_message(msg)
                    break
                if(feedback=='2'):    
                    relatorio = openpyxl.Workbook()
                    print(relatorio.sheetnames)
                    relatorio.create_sheet('relatorio do atendimento')
                    relatorio_page = relatorio['relatorio do atendimento']
                    relatorio_page.append([  data, nome, matricula, 'resolvido', 'bom',])
                    relatorio.save('Relatorio de atendimento.xlsx')
                    msg = EmailMessage()
                    msg['Suject'] = 'Feedback do atendimento do chatbot'
                    msg['from'] = 'pugan chatbot'
                    msg['To'] = email_colaborador 
                    msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                        smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                        smtp.send_message(msg)
                    break
                if(feedback=='3'):    
                    relatorio = openpyxl.Workbook()
                    print(relatorio.sheetnames)
                    relatorio.create_sheet('relatorio do atendimento')
                    relatorio_page = relatorio['relatorio do atendimento']
                    relatorio_page.append([  data, nome, matricula, 'resolvido', 'médio',])
                    relatorio.save('Relatorio de atendimento.xlsx')
                    msg = EmailMessage()
                    msg['Suject'] = 'Feedback do atendimento do chatbot'
                    msg['from'] = 'pugan chatbot'
                    msg['To'] = email_colaborador 
                    msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                        smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                        smtp.send_message(msg)
                    break
                if(feedback=='4'):    
                    relatorio = openpyxl.Workbook()
                    print(relatorio.sheetnames)
                    relatorio.create_sheet('relatorio do atendimento')
                    relatorio_page = relatorio['relatorio do atendimento']
                    relatorio_page.append([  data, nome, matricula, 'resolvido', 'ruim',])
                    relatorio.save('Relatorio de atendimento.xlsx')
                    msg = EmailMessage()
                    msg['Suject'] = 'Feedback do atendimento do chatbot'
                    msg['from'] = 'pugan chatbot'
                    msg['To'] = email_colaborador 
                    msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                        smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                        smtp.send_message(msg)    
                    break
                if(feedback=='5'):    
                    relatorio = openpyxl.Workbook()
                    print(relatorio.sheetnames)
                    relatorio.create_sheet('relatorio do atendimento')
                    relatorio_page = relatorio['relatorio do atendimento']
                    relatorio_page.append([  data, nome, matricula, 'resolvido', 'muito ruim',])
                    relatorio.save('Relatorio de atendimento.xlsx')
                    msg = EmailMessage()
                    msg['Suject'] = 'Feedback do atendimento do chatbot'
                    msg['from'] = 'pugan chatbot'
                    msg['To'] = email_colaborador 
                    msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                    with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                        smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                        smtp.send_message(msg)
                    break
                if(conclusao=='n'):
                    print('É uma pena que não pude lhe ajudar. \n')
                    # Criado e desenvolvido por: Marcelo Pugan
                    conclusao_ruim=input('Prefere [1]Finalizar o atendimento? ou \n[2]recomeçar?\n\n')
                    if(conclusao_ruim == '1'):
                        print('antes de finalizarmos, pode avaliar como foi meu atendimento? \n')
                    feedback = input('[1]muito bom \n[2]bom \n[3]médio \n[4]ruim \n[5]muito ruim \n\n')
                    if(feedback=='1'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='2'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'bom',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='3'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'médio',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='4'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(feedback=='5'):    
                        relatorio = openpyxl.Workbook()
                        print(relatorio.sheetnames)
                        relatorio.create_sheet('relatorio do atendimento')
                        relatorio_page = relatorio['relatorio do atendimento']
                        relatorio_page.append([  data, nome, matricula, 'não resolvido', 'muito ruim',])
                        relatorio.save('Relatorio de atendimento.xlsx')
                        msg = EmailMessage()
                        msg['Suject'] = 'Feedback do atendimento do chatbot'
                        msg['from'] = 'pugan chatbot'
                        msg['To'] = email_colaborador 
                        msg.set_content('esse é o e-mail de confirmação do atendimento ocorrido. ')
                        with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
                            smtp.login(EMAIL_ADDRESS,EMAIL_PASSWORD)
                            smtp.send_message(msg)
                        break
                    if(conclusao_ruim == '2'):
                        print('Recomeçando')
    






     # Criado e desenvolvido por: Marcelo Pugan
