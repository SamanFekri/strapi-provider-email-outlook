const axios = require("axios");
const fs = require("fs");


module.exports = {
  init: (providerOptions = {}, settings = {}) => {
    // get access data
    module.exports.getAcceesData(providerOptions).then((accessData) => {
      module.exports.accessData = accessData;
      console.log("✔ Email provider Get token from Microsoft Graph API");
    }).catch((err) => {
      throw new Error(err);
    });

    return {
      send: async options => {
        // if access data is not set, throw an error
        if (!module.exports.accessData) {
          throw new Error("Access data is not set.");
        }

        // if access token is expired, get a new one
        if (module.exports.accessData.expires_on - Math.floor(Date.now() / 1000) <= 0) {
          
          module.exports.accessData = await module.exports.getAcceesData(providerOptions);
          console.log(module.exports.accessData);
          // if the response is not successful, throw an error
          if (!module.exports.accessData) {
            throw new Error("Access data is not set.");
          }
          console.log("✔ Email provider Access token expired, getting a new one");
        }

        let from = options.from || settings.defaultFrom;
        let to = options.to || settings.defaultTo;
        let token = module.exports.accessData.access_token;
        let token_type = module.exports.accessData.token_type;
        let url = `${module.exports.accessData.resource}/v1.0/users/${from}/sendMail`;
        let cc = options.cc;
        let bcc = options.bcc;

        // make from string format to array suitable for graph api
        to = to ? to.split(",").map((item) => { return {emailAddress: {address: item.trim()}}}): [];
        cc = cc ? cc.split(",").map((item) => { return {emailAddress: {address: item.trim()}}}): [];
        bcc = bcc ? bcc.split(",").map((item) => { return {emailAddress: {address: item.trim()}}}): [];

        let attachments = options.attachments ? options.attachments : [];
        // if attachment is array of file, loop through it and convert it to base64
        if (attachments && Array.isArray(attachments)) {
          attachments = attachments.map((item) => {
            let file = fs.readFileSync(item.path);
            return {
              "@odata.type": "#microsoft.graph.fileAttachment",
              name: item.name,
              contentBytes: file.toString("base64"),
              contentType: item.type,
            };
          });
        }

        return axios({
          method: "POST",
          url,
          data: {
            message: {
              subject: options.subject,
              body: {
                contentType: "HTML",
                content: options.html || options.text,
              },
              toRecipients: to,
              ccRecipients: cc,
              bccRecipients: bcc,
              attachments: attachments,
            },
          },
          headers: {
            Authorization: `${token_type} ${token}`,
            "Content-Type": "application/json",
          },
        });
      },
    };
  },

  getAcceesData: async (providerOptions = {}) => {
    if (!providerOptions.tenant_id) {
      throw new Error("Missing tenant_id.");
    }
    let url = `https://login.microsoftonline.com/${providerOptions.tenant_id}/oauth2/token`;
    let params = new URLSearchParams({
      grant_type: providerOptions.grant_type,
      client_id: providerOptions.client_id,
      client_secret: providerOptions.client_secret,
      resource: providerOptions.resource,
    })
    let res = await axios({
      method: "POST",
      url,
      data: params,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
    });
    // if the response is successful, throw an error
    if (res.status >= 400) {
      throw new Error(res.statusText);
    }
    return res.data;
  },

  accessData: null,
};