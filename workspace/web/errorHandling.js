class LoginError {
    constructor(issue) {
        this.type = 'login';
        this.msg = `The SharePoint ${issue} you entered is incorrect.`;
    }
}

class MondayError {
    constructor(issue) {
        this.type = 'monday';

        if (typeof issue === 'object' && issue[1] === 'Not Found') {
            this.msg = `Could not find the ${issue[0]} board. Please try another.`;
        }
    }
}