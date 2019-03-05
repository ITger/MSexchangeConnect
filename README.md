# MSexchangeConnect
Sample java project on how to connect to a mail account on outlook.com and get some data from it. 

In case you want to use your private corporation ms exchange server it's possible that you need to import some certificates:
1. you should login to your outlook accont using web browser.
2. the browser ask to accept certificates.
3. eksport from the browser proper certificates to a file
4. run:
sudo keytool -import -alias YOUR-CERT-ALIAS-NAME -file YOUR-CERTIFICATE-FILE-NAME.crt -keystore /FULL_PATH_TO_YOUR-JAVA_/security/cacerts 
