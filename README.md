# Strapi Provider Email Outlook

The Strapi Provider Email Outlook is a plugin designed to integrate Outlook email services seamlessly into your Strapi project. This plugin enables you to send emails directly from your Strapi application using Outlook's email service.

## Instalation

- Install the code from npm

``` bash
# if you use npm
npm install strapi-provider-email-outlook
# if you use yarn
yarn add strapi-provider-email-outlook
```

- Add this code in your `./config/plugin.js` of your strapi

``` js
module.exports = ({ env }) => ({
    // ...
    email: {
      config: {
        provider: 'strapi-provider-email-outlook',
        providerOptions: {
          client_id: env('OUTLOOK_CLIENT_ID'),
          client_secret: env('OUTLOOK_CLIENT_SECRET'),
          tenant_id: env('OUTLOOK_TENANT_ID'),
          resource: env('OUTLOOK_RESOURCE', 'https://graph.microsoft.com'),
          grant_type: env('OUTLOOK_GRANT_TYPE', 'client_credentials'),
        },
        settings: {
          defaultFrom: 'example@example.com',
          defaultReplyTo: 'example@example.com',
        },
      },
    }
    // ...
  });
```

### Environment Variables

``` .env
OUTLOOK_CLIENT_ID=<client-id>
OUTLOOK_CLIENT_SECRET=<client-secret>
OUTLOOK_TENANT_ID=<tenant-id>
OUTLOOK_GRANT_TYPE=client_credentials
OUTLOOK_RESOURCE=https://graph.microsoft.com
```

## Road Map

- *Multiple Participants Support:* Enable sending emails to multiple recipients.
- *Reply*: Implement functionality to reply emails
- *Forward*: Implement functionality to forward emails

## Fork and Pull Requests

Feel free to fork this repository and make any changes or enhancements. If you've added features or fixed bugs, submit a pull request. Contributions are welcome!