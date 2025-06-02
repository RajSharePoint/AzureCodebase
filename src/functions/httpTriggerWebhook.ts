import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
// import { ServiceBusClient } from "@azure/service-bus"; // No longer directly used here
import { AzureServiceBusService, IServiceBusService } from '../serviceBusService'; // Import the new service

// Define interfaces for SharePoint webhook notifications
// Place these at the top of your file or in a dedicated types.ts
interface SharePointNotification {
    subscriptionId: string;
    clientState?: string; // Optional, if you set it during subscription
    expirationDateTime: string;
    resource: string;     // ID of the list (or other resource)
    tenantId: string;
    siteUrl: string;      // Server-relative URL of the site (e.g., /sites/MySite)
    webId: string;
    changeType?: string;  // This is the field in question. Mark as optional for now.
}

interface SharePointNotificationPayload {
    value: SharePointNotification[];
}

export async function httpTriggerWebhook(
    request: HttpRequest,
    context: InvocationContext
): Promise<HttpResponseInit> {
    context.log(`Http function processed request for url "${request.url}"`);

    const name = request.query.get('name') || await request.text() || 'world';

    return { body: `Hello, ${name}!` };
};
export async function validateTokenHandler(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    // This is the function registered with Azure Functions.
    // It calls the core logic, allowing the core logic to be tested with a mock.
    return actualValidateTokenHandler(request, context, undefined);
}

// This function contains the core logic and can be called directly by tests with a mock service.
export async function actualValidateTokenHandler(
    request: HttpRequest,
    context: InvocationContext,
    serviceBusServiceOverride?: IServiceBusService // Allow injecting a mock for testing
): Promise<HttpResponseInit> {
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
                let parsedPayload: SharePointNotificationPayload;
                try {
                    // SharePoint sends notifications as JSON.
                    parsedPayload = await request.json() as SharePointNotificationPayload;
                } catch (e) {
                    context.error("Error parsing notification payload as JSON:", e instanceof Error ? e.message : String(e));
                    // Log the raw text if JSON parsing fails, for debugging
                    try {
                        const rawPayloadForError = await request.text(); // Attempt to read as text for logging
                        context.log("Raw notification payload (on JSON parse error):", rawPayloadForError);
                    } catch (textError) {
                        context.error("Could not even read raw payload as text after JSON parse error:", textError);
                    }
                    return {
                        status: 400, // Bad Request
                        body: "Invalid JSON payload for notification."
                    };
                }

                if (!parsedPayload) {
                    context.log("Notification payload is empty or not in the expected format (missing 'value' array).");
                    // Log the parsed payload if it exists but is not as expected
                    if(parsedPayload) context.log("Received payload structure (malformed?):", JSON.stringify(parsedPayload));
                    return {
                        status: 400, // Bad Request
                        body: "Notification payload is empty or malformed."
                    };
                }

                // CRUCIAL LOGGING: Log the entire parsed payload to inspect its structure.
                context.log(`Received notification(s). Full payload: ${JSON.stringify(parsedPayload, null, 2)}`);

                // Attempt to send notifications to Azure Service Bus
                const connectionString = process.env.SERVICE_BUS_CONNECTION_STRING;
                const queueName = process.env.SERVICE_BUS_QUEUE_NAME;

                if (connectionString && queueName) {
                    const sbService = serviceBusServiceOverride || new AzureServiceBusService(connectionString, queueName);
                    const messagesToSend = [];

                    for (const notification of parsedPayload.value) {
                        context.log(`Processing individual notification for Service Bus: SubscriptionId: ${notification.subscriptionId}, Resource (ListID): ${notification.resource}, SiteUrl: ${notification.siteUrl}`);

                        if (notification.changeType) {
                            context.log(`  ChangeType found: ${notification.changeType}`);
                            // Now you can use notification.changeType to determine the event
                            // e.g., if (notification.changeType === 'added') { /* ... */ }
                        } else {
                            context.warn(`  WARNING: 'changeType' field is missing in this notification object: ${JSON.stringify(notification, null, 2)}`);
                            // If changeType is consistently missing, you'll need to investigate your webhook subscription
                            // or the specific SharePoint event source.
                        }

                        messagesToSend.push({
                            body: notification, // Send the entire notification object
                            contentType: "application/json",
                            // You could add a messageId or other properties if needed
                            // messageId: `${notification.subscriptionId}_${new Date().toISOString()}`
                        });
                    }

                    if (messagesToSend.length > 0) {
                        try {
                            await sbService.sendMessages(messagesToSend, context);
                            // Logging related to successful send is now within AzureServiceBusService
                        } catch (sbError) {
                            // Error is already logged by AzureServiceBusService.
                            // Depending on your requirements, you might want to implement a retry mechanism
                            // or dead-lettering for failed messages. For now, we log and continue.
                        }
                    }
                } else {
                    context.warn("Service Bus environment variables (SERVICE_BUS_CONNECTION_STRING or SERVICE_BUS_QUEUE_NAME) are not set. Skipping message queuing.");
                }
                // SharePoint expects a quick response (within 5 seconds typically for notifications too).
                // A 202 Accepted is good if processing is asynchronous.
                return { status: 202, body: "Notification acknowledged and queued for processing." };
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
        context.error(`Error processing actualValidateTokenHandler: ${errorMessage}`);
        return {
            status: 500,
            body: "An internal error occurred while processing the request."
        };
    }
}

app.http('tokenEndpoint', {
    methods: ['GET','POST'],
    authLevel: 'anonymous',
    handler: validateTokenHandler // The registered handler remains validateTokenHandler
});


app.http('httpTriggerWebhook', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: httpTriggerWebhook
});
