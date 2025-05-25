import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

export async function httpTriggerWebhook(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log(`Http function processed request for url "${request.url}"`);

    const name = request.query.get('name') || await request.text() || 'world';

    return { body: `Hello, ${name}!` };
};
export async function validateTokenHandler(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log(`Validate Token function processed request for url "${request.url}", method: ${request.method}`);

    try {
        // SharePoint webhook validation: POST request with 'validationtoken' in query string.
        // SharePoint webhook notification: POST request with payload in body.
        if (request.method === 'POST') {
            const validationTokenFromQuery = request.query.get('validationtoken');

            if (validationTokenFromQuery) {
                // This is the validation request from SharePoint.
                context.log(`SharePoint validation POST request. Token from query: "${validationTokenFromQuery}". Responding with 200 OK.`);
                return {
                    status: 200, // CRITICAL: SharePoint expects 200 OK for validation.
                    headers: { "Content-Type": "text/plain" },
                    body: validationTokenFromQuery // Return the exact token from the query.
                };
            } else {
                // This is a notification request from SharePoint.
                // The payload is in the request body.
                const notificationPayload = await request.text(); // Or request.json() if notifications are JSON.
                
                if (!notificationPayload) {
                    context.log("Notification POST request body is missing or empty.");
                    return {
                        status: 400, // Bad Request
                        body: "Notification body is missing or empty."
                    };
                }

                context.log(`SharePoint notification POST request. Body content: "${notificationPayload}".`);
                // TODO: Process the notification payload (notificationPayload) here.
                // For example, queue it for asynchronous processing, save to a database, etc.
                // SharePoint expects a quick response (within 5 seconds typically for notifications too).
                // A 202 Accepted is good if processing is asynchronous.
                // A 200 OK is fine if processing is synchronous and fast.
                return { status: 202, body: "Notification acknowledged" }; 
            }
        } else if (request.method === 'GET') {
            // This endpoint is configured to accept GET, but SharePoint webhooks (validation/notification) use POST.
            // If a GET request comes with a 'validationtoken', it might be a misconfiguration or a different client.
            const validationTokenFromQueryForGET = request.query.get('validationtoken');
            if (validationTokenFromQueryForGET) {
                context.log(`Received GET request with 'validationtoken': "${validationTokenFromQueryForGET}". SharePoint webhook validation uses POST. Responding with token for compatibility/testing if needed.`);
                // While docs say POST, if a GET with token comes, responding might help debug or cover edge cases.
                return {
                    status: 200,
                    headers: { "Content-Type": "text/plain" },
                    body: validationTokenFromQueryForGET
                };
            }
            
            context.log("Received GET request. This endpoint is primarily for SharePoint webhook POST requests. No 'validationtoken' found in query for GET.");
            return {
                status: 405, // Method Not Allowed for this specific SharePoint webhook interaction
                body: "SharePoint webhooks use POST. Validation: POST with 'validationtoken' in query. Notifications: POST with payload in body."
            };
        }

        context.log(`Unsupported HTTP method: ${request.method}. This endpoint expects POST for SharePoint webhooks or GET for specific tests.`);
        return { status: 405, body: `Method ${request.method} not allowed or not handled.` };

    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        context.error(`Error processing validateTokenHandler: ${errorMessage}`);
        return {
            status: 500,
            body: "An internal error occurred while processing the request."
        };
    }
}

app.http('tokenEndpoint', {
    methods: ['GET','POST'],
    authLevel: 'anonymous',
    handler: validateTokenHandler
});


app.http('httpTriggerWebhook', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: httpTriggerWebhook
});
