const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');
const https = require('https');
const multer = require('multer');
const FormData = require('form-data');

const app = express();
const PORT = process.env.PORT || 3000;

// Create an HTTPS agent that does not validate certificates
const agent = new https.Agent({ rejectUnauthorized: false });

// Enable CORS for all origins
app.use(cors());

// Configure Multer for file uploads
const upload = multer();

app.post('/api/proxy/adobeSign', upload.single('File'), async (req, res) => {
    //Get TransientID
    const proxyUrl = "https://secure.na4.adobesign.com/api/rest/v6/transientDocuments";

    try {
        const form = new FormData();
        form.append('File', req.file.buffer, req.file.originalname);

        const response = await fetch(proxyUrl, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer 3AAABLblqZhD7WMt-0-z8I2BWXe6FaZAN68Y3piUGB8uW_1_LVBoo3IQalQFcF4Zz7HO6vmE2ji-HymCHZBbvSOp_TAy-5h0-`,
                ...form.getHeaders(), // Get headers generated by the FormData instance
            },
            body: form,
            agent: agent
        });

        const data = await response.json();

        if (!data.transientDocumentId) {
            throw new Error('No transientDocumentId returned from Adobe Sign');
        }

        const transientDocumentId = data.transientDocumentId;
        console.log(transientDocumentId);
        
        const fileNameWithExtension = req.file.originalname;
        const OGfileName = fileNameWithExtension.substring(0, fileNameWithExtension.lastIndexOf('.')) || fileNameWithExtension;

        //Get AgreementID
        const createAgreementUrl = "https://secure.na4.adobesign.com/api/rest/v6/agreements";
        const agreementBody = JSON.stringify({
            "fileInfos": [
                {
                    "transientDocumentId": transientDocumentId
                }
            ],
            "name": OGfileName,
            "signatureType": "ESIGN",
            "state": "DRAFT"
        });

        const agreementResponse = await fetch(createAgreementUrl, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer 3AAABLblqZhD7WMt-0-z8I2BWXe6FaZAN68Y3piUGB8uW_1_LVBoo3IQalQFcF4Zz7HO6vmE2ji-HymCHZBbvSOp_TAy-5h0-`,
                'Content-Type': 'application/json'
            },
            body: agreementBody,
            agent: agent
        });

        const agreementData = await agreementResponse.json();

        if (!agreementData.id) {
            throw new Error('No agreementId returned from Adobe Sign');
        }

        const agreementId = agreementData.id;
        console.log(agreementId);

        //Get Agreement Views
        const viewUrl = `https://secure.na4.adobesign.com/api/rest/v6/agreements/${agreementId}/views`;
        const viewBody = JSON.stringify({
            "name": "ALL"
        });

        const viewResponse = await fetch(viewUrl, {
            method: 'POST',
            headers: {
                'Authorization': `Bearer 3AAABLblqZhD7WMt-0-z8I2BWXe6FaZAN68Y3piUGB8uW_1_LVBoo3IQalQFcF4Zz7HO6vmE2ji-HymCHZBbvSOp_TAy-5h0-`,
                'Content-Type': 'application/json'
            },
            body: viewBody,
            agent: agent
        });

        const viewData = await viewResponse.json();

        res.status(viewResponse.status).json(viewData);

    } catch (error) {
        console.error('Error forwarding request to Adobe Sign:', error);
        res.status(500).json({ error: 'Error communicating with Adobe Sign' });
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Proxy server running on port ${PORT}`);
});