import { getEmployeeProfile } from './northwindDataService.js';
import { EMAIL_DOMAIN } from './constants.js';

export async function getUserProfile(userId) {

    const employeeProfile = await getEmployeeProfile(userId);

    return {
        displayName: `${employeeProfile.FirstName} ${employeeProfile.LastName}`,
        mail: `${employeeProfile.FirstName}@${EMAIL_DOMAIN}`,
        photo: employeeProfile.Photo.substring(104), // Trim Northwind-specific junk
        jobTitle: employeeProfile.Title
    }

}

