
"""
@author: Arnav Jain
"""

# Importing Modules
import smtplib
from email.message import EmailMessage
import pandas as pd

# Reading the xlsx file
file = pd.read_excel('Email Excel File (xlsx)')

## Reading each column
names = file.loc[:, "Name Column Header"].values
mails = file.loc[:, "Email ID Column Header"].values

for x in range(len(names)):
    print(f'{names[x]}: {mails[x]}')

# Storing the credentials

# ------------------------------------------------------------------------------------------
## Make sure to enable 'Allow Less Secure Apps Access' in the senders google account settings
# ------------------------------------------------------------------------------------------

address = 'Sender Email Address'
password = 'Sender Email Password'

# Creating the Loop
for x in range(len(names)):

    # Designating the Recievers, Senders and First Names
    msg = EmailMessage()
    msg['subject'] = 'Invitation for Marketing Workshop'
    msg['From'] = address
    msg['To'] = mails[x]
    fname = str(names[x].split()[0]) # Only if wanted or necessary
    msg.add_alternative(

    # Adding the Body of the mail through HTML

    f"""\
    Hello {fname}, \n
    <!DOCTYPE html>
    <html>
        <body>
            <p>
                Lorem ipsum dolor sit amet consectetur adipisicing elit. Maxime mollitia,
                molestiae quas vel sint commodi repudiandae consequuntur voluptatum laborum
                numquam blanditiis harum quisquam eius sed odit fugiat iusto fuga praesentium
                optio, eaque rerum! Provident similique accusantium nemo autem. Veritatis
                obcaecati tenetur iure eius earum ut molestias architecto voluptate aliquam
                nihil, eveniet aliquid culpa officia aut! Impedit sit sunt quaerat, odit,
                tenetur error, harum nesciunt ipsum debitis quas aliquid. Reprehenderit,
                quia. Quo neque error repudiandae fuga? Ipsa laudantium molestias eos 
                sapiente officiis modi at sunt excepturi expedita sint? Sed quibusdam
                recusandae alias error harum maxime adipisci amet laborum. Perspiciatis 
                minima nesciunt dolorem! Officiis iure rerum voluptates a cumque velit 
                quibusdam sed amet tempora. Sit laborum ab, eius fugit doloribus tenetur 
                fugiat, temporibus enim commodi iusto libero magni deleniti quod quam 
                consequuntur! Commodi minima excepturi repudiandae velit hic maxime
                doloremque. Quaerat provident commodi consectetur veniam similique ad 
                earum omnis ipsum saepe, voluptas, hic voluptates pariatur est explicabo 
                fugiat, dolorum eligendi quam cupiditate excepturi mollitia maiores labore 
                suscipit quas? Nulla, placeat. Voluptatem quaerat non architecto ab laudantium
                modi minima sunt esse temporibus sint culpa, recusandae aliquam numquam 
                totam ratione voluptas quod exercitationem fuga. Possimus quis earum veniam 
                quasi aliquam eligendi, placeat qui corporis!
            </p>
        </body>
    </html>
    """, subtype='html')
    
    # Sending The Mail
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(address, password)
        smtp.send_message(msg)
        print("Sent to:", names[x], "Row Number:", x)
