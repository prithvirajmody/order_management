from twilio.rest import Client

def send_message(message_body):
    # Your Twilio account SID and Auth Token
    account_sid = 'ACa9a6a6dce084ae1407612116d612f0a8'
    auth_token = '72ffb22078ebb91b124e5b861fe90e27'

    # Create a Twilio client
    client = Client(account_sid, auth_token)

    # Your Twilio phone number and the recipient's phone number
    twilio_phone_number = '+19087606125'
    recipient_phone_number = '+919004372646'  # Replace with the recipient's actual phone number

    # Send the message
    message = client.messages.create(
        body=message_body,
        from_=twilio_phone_number,
        to=recipient_phone_number
    )

#send_message('test')