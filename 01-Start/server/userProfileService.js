export async function getUserProfile() {

    return {
        displayName: "Angel Brown",
        mail: "angel@northwind.com",
        company: process.env.COMPANY_NAME,
        jobTitle: "Team lead"
    }

}

