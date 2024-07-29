from flask import Flask, render_template, request, redirect, url_for
import asyncio
import configparser
from msgraph.generated.models.o_data_errors.o_data_error import ODataError
from graph import Graph

app = Flask(__name__)

# Load settings
config = configparser.ConfigParser()
config.read(['config.cfg', 'config.dev.cfg'])
azure_settings = config['azure']

graph = Graph(azure_settings)

@app.route('/')
def index():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    user = loop.run_until_complete(graph.get_user())
    if user:
        user_email = user.mail or user.user_principal_name
        return render_template('index.html', user=user.display_name, email=user_email)
    return 'User not found', 404

@app.route('/list-inbox')
def list_inbox():
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        message_page = loop.run_until_complete(graph.get_inbox())
        messages = []
        if message_page and message_page.value:
            for message in message_page.value:
                messages.append({
                    'subject': message.subject,
                    'from': message.from_.email_address.name if message.from_ and message.from_.email_address else 'NONE',
                    'status': 'Read' if message.is_read else 'Unread',
                    'received': message.received_date_time
                })

        return render_template('inbox.html', messages=messages)
    except ODataError as odata_error:
        error_message = f'{odata_error.error.code} - {odata_error.error.message}' if odata_error.error else 'Unknown error'
        return error_message, 500

@app.route('/getAToken')
def get_a_token():
    code = request.args.get('code')
    azure_settings['authorization_code'] = code
    graph._acquire_token()
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)
