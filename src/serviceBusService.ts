import { ServiceBusClient, ServiceBusMessage } from "@azure/service-bus";
import { InvocationContext } from "@azure/functions";

export interface IServiceBusService {
    sendMessages(messages: ServiceBusMessage[], context: InvocationContext): Promise<void>;
}

export class AzureServiceBusService implements IServiceBusService {
    private connectionString: string;
    private queueName: string;

    constructor(connectionString: string, queueName: string) {
        if (!connectionString) {
            throw new Error("Service Bus connection string is required.");
        }
        if (!queueName) {
            throw new Error("Service Bus queue name is required.");
        }
        this.connectionString = connectionString;
        this.queueName = queueName;
    }

    async sendMessages(messages: ServiceBusMessage[], context: InvocationContext): Promise<void> {
        const sbClient = new ServiceBusClient(this.connectionString);
        const sender = sbClient.createSender(this.queueName);

        try {
            context.log(`Attempting to send ${messages.length} message(s) to Service Bus queue '${this.queueName}'.`);
            await sender.sendMessages(messages);
            context.log(`${messages.length} message(s) successfully sent to Service Bus queue '${this.queueName}'.`);
        } catch (sbError) {
            const sbErrorMessage = sbError instanceof Error ? sbError.message : String(sbError);
            context.error(`Failed to send messages to Service Bus queue '${this.queueName}': ${sbErrorMessage}`);
            // Re-throw the error so the caller can be aware and potentially handle it (e.g., retry, dead-letter)
            throw sbError;
        } finally {
            try {
                await sender.close();
                await sbClient.close();
                context.log("Service Bus sender and client closed after attempting to send messages.");
            } catch (closeError) {
                // Log error during close but don't let it overshadow a potential send error
                context.error(`Error closing Service Bus client/sender: ${closeError instanceof Error ? closeError.message : String(closeError)}`);
            }
        }
    }
}
