import { getEmployeeProfile } from './northwindData.js';

export async function getUserProfile() {

    const employeeProfile = await getEmployeeProfile(1);

    return {
        displayName: `${employeeProfile.FirstName} ${employeeProfile.LastName}`,
        mail: `${employeeProfile.FirstName}@northwindtraders.com`,
        photo: employeeProfile.Photo.substring(104), // Trim Northwind-specific junk
        company: process.env.COMPANY_NAME,
        jobTitle: employeeProfile.Title
    }

}

