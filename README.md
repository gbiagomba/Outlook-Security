# Outlook-Security
These scripts are designed to allow users to forward potentially malicious emails to their IT Security team and append the "encrypt" tag to the subject line

## How to import these scripts
Follow the instructions in any of the below links then add the code you wish to use in the dev console (once you get to it)

```
https://www.howto-outlook.com/howto/macrobutton.htm
https://www.datanumen.com/blogs/run-vba-code-outlook/
https://docs.microsoft.com/en-us/office/vba/outlook/how-to/using-visual-basic-to-customize-outlook-forms/about-using-vbscript-in-outlook
```

## Manifesto
**outlook-encrypt.vbs**: This script is to add the "encrypt" tag to the subject line, it doesnt actually try to encrypt your email. Here we are summing you are running something like proofpoint, and that solution will know to turn on email encryption.

**outlook-security.vbs**: This script is going to forward any email you have selected as an attachment and it will automatically delete it for you. This is great if you think the email is malicious and you want to forward it to your IT Security team. 