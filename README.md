# CordovaMSPOC
simple POC for testing Microsoft Azure &amp; Graph APIs from Cordova.

This is based on the example [Azure AD Cordova getting started](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-devquickstarts-cordova) which used the old Azure URLs. This allows you to swap between them to compare my **changing the var useV2 in www/index.js**. You will need to change **redirectUri** and **clientId** to match your site. It is still **VERY primative** as it is mainly just for trying some quick calls to confirm they work and see what is retuned in console before creating ionic version. 

Look down around line 80 to swap in the URLs you want called.

I start with a script that runs the following commands

**call cordova clean**

**call cordova emulate android**

**adb logcat**
