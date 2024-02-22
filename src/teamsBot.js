const { TeamsActivityHandler, TurnContext } = require('botbuilder');
const fs = require('fs');
const path = require('path');
const cron = require('node-cron');

// Archivo JSON para almacenar los IDs de conversación
const conversationIdsFile = 'conversationIds.json';

// Carga los IDs de conversación existentes desde el archivo JSON
let conversations = loadConversationsFromJson();

class TeamsBot extends TeamsActivityHandler {
    constructor() {
        super();

        // Maneja eventos de miembros del equipo para obtener el ID de la conversación
        this.onMembersAdded((context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (let member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    // Obtiene el ID de la conversación
                    const conversationId = TurnContext.getConversationReference(context.activity).conversation.id;

                    // Almacena el contexto de la conversación
                    storeConversationContext(conversationId, context);

                    console.log(`Conversación agregada. ID: ${conversationId}`);
                }
            }
            return next();
        });

        // Maneja mensajes de usuarios
        this.onMessage((context, next) => {
            // Obtén el ID de conversación
            const conversationId = TurnContext.getConversationReference(context.activity).conversation.id;

            // Almacena el contexto de la conversación
            storeConversationContext(conversationId, context);

            console.log(`Mensaje recibido en la conversación. ID: ${conversationId}`);

            // Procesa el mensaje
            this.processMessage(context);

            return next();
        });
    }

    processMessage(context) {
        // Agrega tu lógica de procesamiento de mensajes aquí
        const text = context.activity.text;
        if (text === 'test') {
            console.log('Mensaje recibido para enviar un mensaje proactivo.');

            // Envía un mensaje proactivo
            const message = 'Hola, este es un mensaje proactivo.';
            this.sendProactiveMessage(context, message);
        }
    }

    sendProactiveMessage(context, message) {
        const conversationId = TurnContext.getConversationReference(context.activity).conversation.id;
        const storedContext = getConversationContext(conversationId);

        if (storedContext) {
            // Envía el mensaje proactivo
            console.log(`Enviando mensaje proactivo a la conversación. ID: ${conversationId}`);
            storedContext.sendActivity({ type: 'message', text: message })
                .catch(error => console.error('Error al enviar mensaje proactivo:', error));
        } else {
            console.error('Contexto de conversación no encontrado.');
        }
    }
}

// Almacena el contexto de la conversación
const storeConversationContext = (conversationId, context) => {
    conversations[conversationId] = context;
    // Guarda los IDs de conversación actualizados en el archivo JSON
    saveConversationsToJson(conversations);
};

// Obtiene el contexto de la conversación
const getConversationContext = (conversationId) => {
    return conversations[conversationId];
};

// Carga los IDs de conversación desde el archivo JSON
function loadConversationsFromJson() {
    try {
        const jsonData = fs.readFileSync(conversationIdsFile, 'utf8');
        return JSON.parse(jsonData) || {};
    } catch (error) {
        console.error('Error al cargar el archivo JSON:', error);
        return {};
    }
}

// Guarda los IDs de conversación en el archivo JSON
function saveConversationsToJson(conversations) {
    try {
        if (!fs.existsSync(conversationIdsFile)) {
            const jsonData = JSON.stringify(conversations, null, 2);
            fs.writeFileSync(conversationIdsFile, jsonData, 'utf8');
        }
    } catch (error) {
        console.error('Error al guardar en el archivo JSON:', error);
    }
}

module.exports.TeamsBot = TeamsBot;
