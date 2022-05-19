import email, imaplib, discord
from discord.ext import commands

service = imaplib.IMAP4_SSL('outlook.office365.com')
service.login('email', 'password')

# Email Folder
def email_select(name='Inbox'):
    try:
        service.select(name)
    except:
        raise 'Invalid Folder Name'

# Get Current Email
def get_num_email():
    num = service.search(None, 'ALL')[1][0].split()
    num.reverse()
    return num

# Helper Function to get from email, to email, BCC, subject 
# and content
def from_email(msg):
    return msg.get('From')

def to_email(msg):
    return msg.get('To')

def bcc(msg):
    return msg.get('BCC')

def subject(msg):
    return msg.get('Subject')

def content(msg):
    output = []
    for part in msg.walk():
        if part.get_content_type() == 'text/plain':
            output.append(part.get_payload())
    
    return ''.join(output)

# Message Function that will return specified amounts of email
def get_email(num=0):
    output = [0 for _ in range(num)]
    total_num = get_num_email()


    for i in total_num:
        if not num:
            break
    
        data = service.fetch(i, '(RFC822)')[1]
        message = email.message_from_bytes(data[0][1])

        fr_email = from_email(message)
        t_email = to_email(message)
        bc = bcc(message)
        sub = subject(message)
        con = content(message)

        current_msg = {
            'from': fr_email,
            'to': t_email,
            'bcc': bc,
            'subject': sub,
            'content': con,
        }

        output[num-1] = current_msg

        num -= 1
    
    return output

# Get Latest Email
def get_latest_email():
    return get_email(1)[0]


bot = commands.Bot(command_prefix='!')

@bot.command(name='test')
async def test(ctx):
    output = f'```Bro```'
    await ctx.send(output)


@bot.command(name='new')
async def get_new(ctx):
    email_select()
    msg = get_latest_email()
    output = f'```From: {msg["from"]}\nTo: {msg["to"]}\nBCC: {msg["bcc"]}\nSubject: {msg["subject"]}\nContent:\n{msg["content"]}```'
    await ctx.send(output)

bot.run('OAUTH')
# service.logout()
