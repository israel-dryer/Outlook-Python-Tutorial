"""
  COPY AND PASTE CODE FOR YOUTUBE TUTORIAL
  
  Video Title: Python Outlook Email - Attachments - Learn how to control Microsoft Outlook using Python
  Video URL: https://youtu.be/omDnG4vO6Wc
  Last Modified: 7/30/2020
"""

html_body = """
    <div>
        <h1 style="font-family: 'Lucida Handwriting'; font-size: 56; font-weight: bold; color: #9eac9c;"> Happy Birthday!! </h1>
        <span style="font-family: 'Lucida Sans'; font-size: 28; color: #8d395c;"> Wishing you all the best on your birthday!! </span>
    </div><br>
    <div>
        <img src="cid:cake-img" width=50%>
    </div>
    """

# code for changing the content id of the image
image.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "cake-img")
