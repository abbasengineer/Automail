import pandas as pd 
import smtplib
my_name = "Abeeza Sheineer"
my_email = "abeezasheineer@gmail.com"
excelFile = pd.read_excel("emailList.xlsx")

emails = excelFile['Email']
names = excelFile['Name']
products = excelFile['Product']
likeness = excelFile['Likeness']

#print("emails: ", emails)
#print("names: ", names)
#print("prod: ", products)
#print("likeness: ", likeness)


server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.ehlo()
server.login(my_email, "Abbfat521!!")

for idx in range(len(emails)):
  name = names[idx]
  email = emails[idx]
  prod = products[idx]
  like = likeness[idx]

  #print(name, email, prod, like)

  bodyMessage = (
    "From: {0} <{1}>\n"
    "To: {2} <{3}>\n"
    "Subject: {4}\n\n"
    "Hello {2}!\n"
    "   My name is {0}, I just wanted to say I am a huge fan and tell you that I have been really enjoying the {4} you have, because {5}."
    "I always recommend your products to my family and friends and will continue to promote the best brand that is {2}! \n"
    "I know some folks email to complain and it can be tough responding to those, so I thought I would add a little sunshine.\n"
    "As someone whos been a loyal customer for years, you are doing fantastic work! \n\n"
    "I am proud to encourage others to invest in your quality products.\n"
    "I would love to try more of your products and would appreciate any samples or coupons you would send my way.\n"
    "Thank you so much! And keep up the amazing work\n"
    "Best,\n\n"
    "{0}".format(my_name, my_email, name, email, prod, like)
  )

  try:
    server.sendmail(my_email, [email], bodyMessage)
    print('Email to {} successfully sent \n'.format(email))
  except Exception as e:
    print('Email to {} could not be sent, because {} \n\n'.format(email, str(e)))

server.close()

