## CalendarReady SPFx Web Part
### Overview
This is a web part built using the new SharePoint Framework. It pulls the current calendar appointments and when a user selects one of those appointments, the web part will pull documents "related" to that meeting. The logic which defines "related" is still being designed but it will lean heavily on the insights layer from the Microsoft Graph.

### Warning!
This code uses some trickery to leverage the Microsoft Graph to get the calendar entries. I am using PostMan to grab a bearer token and passing that in with the httpGet request. Needless to say, this isn't a long term approach. For more information about how I use PostMan to get the token, see this link: http://www.chrisjohnson.io/2015/11/06/simplifying-office-365-unified-api-calls-with-postman-and-oauth-2/

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* commonjs components - this allows this package to be reused from other packages.
* dist/* - a single bundle containing the components used for uploading to a cdn pointing a registered Sharepoint webpart library to.
* example/* a test page that hosts all components in this package.

### Build options

gulp nuke - TODO
gulp test - TODO
gulp watch - TODO
gulp build - TODO
gulp deploy - TODO
