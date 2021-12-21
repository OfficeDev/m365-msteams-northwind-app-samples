// Get the user profile from our web service
export async function getUserProfile(clientSideToken) {

    const response = await fetch ("/userProfile?userId=2", {
        "method": "get",
        "headers": {
            "content-type": "application/json"
        },
        "cache": "default"
    });
    if (response.ok) {
        const userProfile = await response.json();
        return userProfile;
    } else {
        const error = await response.json();
        console.log (`ERROR: ${error}`);
        throw (error);
    }
}
