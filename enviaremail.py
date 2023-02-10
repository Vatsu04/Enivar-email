
import win32com.client as win32

#criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

#criar um email
email = outlook.CreateItem(0)

#Confirgurar as informações do seu emial
email.To = "gugatascheri@gmail.com; pythonimpressionador+lira@gmail.com"
email.Subject = "E-mail automatico python"
email.HTMLbody = """ <p>Ola</p>
<p>fdp</p>

<p><i>abs</i></p>
Código Python"""

anexo = "D:\iames - Copy/612.png"
email.Attachments.Add(anexo)
email.Send()
print("Email enviado")