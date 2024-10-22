# email_content.py

def generate_html_email_content(recipient_name):
    """
    Generate dynamic HTML email content based on the recipient's name.
    """
    # Example HTML email content with dynamic recipient name
    html_email_content = f"""
    <html>
    <body>
        <h1 style="color:blue;">Hello, {recipient_name}!</h1>
        <p>This is a test email sent via Selenium automation.</p>
        <p>Here’s an example of <b>bold text</b> and <i>italic text</i>.</p>
        <p>Thank you for being a valued member.</p>
    </body>
    </html>
    """
    return html_email_content
