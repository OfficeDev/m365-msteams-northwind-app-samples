import dotenv from 'dotenv';
import aad from 'azure-ad-jwt';


// please go to the below link to learn more about SaaS offer creation
// https://docs.microsoft.com/en-us/azure/marketplace/partner-center-portal/offer-creation-checklist

const offer = {
    "OfferID": process.env.OFFEER_ID,
    "Name": process.env.OFFER_NAME
};

export async function validateLicense(accessToken) {

    const audience = `api://${process.env.HOSTNAME}/${process.env.CLIENT_ID}`;
    return new Promise ((resolve, reject) => {

        aad.verify(accessToken, { audience: audience }, (err, result) => {
            if (result) {
                // Call the service here
                resolve(true);
            } else {
                throw ({ status: 401, message: 'Invalid access token' });
            }
        });    
    });

}

