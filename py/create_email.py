import pywinauto
import win32com.client
import os

filename ="//wsl.localhost/Ubuntu/home/burtnolej/sambashare/veloxmon/websitepy/output_articles/generated_newsletter.html"

attachments=["C:/Users/burtn/Development/py/velox_logo_medium.svg", \
             "C:/Users/burtn/Development/py/linkedin_normal_small.svg", \
             "C:/Users/burtn/Development/py/linkedin_normal_vsmall.svg", \
             "C:/Users/burtn/Development/py/medium_normal_small.svg", \
             "C:/Users/burtn/Development/py/medium_normal_vsmall.svg", \
             "C:/Users/burtn/Development/py/twitter_normal_small.svg", \
             "C:/Users/burtn/Development/py/twitter_normal_vsmall.svg", \
             "C:/Users/burtn/Development/py/velox_emblem.svg"]

subject="The Velox Build Faster Digest - October 23"
sender="build.faster@veloxfintech.com" 
cc="jon.butler@veloxfintech.com"

#bcc="joachim.lauterbach@fsa.valantic.com; Andy.Browning@fsa.valantic.com; Peter.Holmgren@fsa.valantic.com; Luigi.Marino@fsa.valantic.com; Greg.Cooper@fsa.valantic.com; Ansgar.Bruell@fsa.valantic.com; andy.bennett@rapidaddition.com; mike.powell@rapidaddition.com; deepak.dhayatker@rapidaddition.com; aaron.pryce@rapidaddition.com; nikolaivarma@googlemail.com; firdausbmi@live.com; reena.raichura@glue42.com; nadine@shokubai.io; giles.sarton@ultumus.com; bernie.thurston@ultumus.com; Janet.Richardson@ultumus.com; femi@ultumus.com; mat.knapman@ultumus.com; jane.lemmon@ultumus.com; Dylan.Myburgh@cognizant.com; Aditya.Sadhotra@cognizant.com; lee.harding@thenorthstarr.com; Toby.Babb@HarringtonStarr.com; bradley.cooke@firthrossmartin.com; simonellis@bmlltech.com; steve@vision57.com; p.andrea@gmail.com; jim@interop.io; loreta.bahtchevanova@interop.io"

html=""
with open(filename,"r+") as f:
    for line in f:
        html=html+line

ol = win32com.client.Dispatch("outlook.Application")
ns = ol.GetNamespace("MAPI")
mail = ol.CreateItem(0)
for _attachment in attachments:
    mail.attachments.Add(_attachment, 1, 0)

mail.subject=subject
#mail.bcc=bcc
mail.cc=cc
mail.sender=sender
mail.HTMLBody = html
mail.save()
mail.display()



