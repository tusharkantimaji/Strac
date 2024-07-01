const express = require('express');
const axios = require('axios');
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 4000;

const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET
    }
};

const cca = new msal.ConfidentialClientApplication(config);

app.get('/auth', (req, res) => {
    const authUrl = cca.getAuthCodeUrl({
        scopes: ['Files.Read', 'Files.Read.All', 'User.Read', 'User.Read.All'],
        redirectUri: process.env.REDIRECT_URI
    });

    res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ['Files.Read', 'Files.Read.All', 'User.Read', 'User.Read.All'],
        redirectUri: process.env.REDIRECT_URI
    };

    try {
        const response = await cca.acquireTokenByCode(tokenRequest);
        req.session.accessToken = response.accessToken;
        res.redirect('/files');
    } catch (error) {
        console.error(error);
        res.status(500).send(error);
    }
});

const getAuthenticatedClient = (accessToken) => {
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken);
        }
    });

    return client;
};

app.get('/files', async (req, res) => {
    const accessToken = req.session.accessToken;
    const client = getAuthenticatedClient(accessToken);

    try {
        const files = await client.api('/me/drive/root/children').get();
        res.json(files.value);
    } catch (error) {
        console.error(error);
        res.status(500).send(error);
    }
});

app.get('/download/:itemId', async (req, res) => {
    const accessToken = req.session.accessToken;
    const client = getAuthenticatedClient(accessToken);

    try {
        const item = await client.api(`/me/drive/items/${req.params.itemId}`).get();
        const downloadUrl = item['@microsoft.graph.downloadUrl'];
        const response = await axios.get(downloadUrl, { responseType: 'stream' });

        res.setHeader('Content-Disposition', `attachment; filename=${item.name}`);
        response.data.pipe(res);
    } catch (error) {
        console.error(error);
        res.status(500).send(error);
    }
});

app.get('/permissions/:itemId', async (req, res) => {
    const accessToken = req.session.accessToken;
    const client = getAuthenticatedClient(accessToken);

    try {
        const permissions = await client.api(`/me/drive/items/${req.params.itemId}/permissions`).get();
        res.json(permissions.value);
    } catch (error) {
        console.error(error);
        res.status(500).send(error);
    }
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
