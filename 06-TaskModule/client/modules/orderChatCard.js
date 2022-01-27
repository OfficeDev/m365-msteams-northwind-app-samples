export default
{
    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
    "type": "AdaptiveCard",
    "version": "1.4",
    "refresh": {
        "userIds": [],
        "action": {
            "type": "Action.Execute",
            "verb": "refresh",
            "title": "Refresh",
            "data": {              
            }
        }
    },
    "body": [
        {
            "type": "ColumnSet",
            "columns": [
                {
                    "type": "Column",
                    "width": "stretch",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Chat with ${contact}: Order #${orderId}",
                            "horizontalAlignment": "left",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                              {
                                "title": "Sales rep:",
                                "value": "${salesRepName}"
                              },
                              {
                                "title": "Sales rep manager:",
                                "value": "${salesRepManagerName}"
                              } 
                                     ]
                            }
                        ]
                }
            ]
        }
      
    ],
    "actions": [
       
        {
            "type": "Action.OpenUrl",
            "title": "Chat with sales rep team",
            "id": "chatWithUser",
            "url": "https://teams.microsoft.com/l/chat/0/0?users=${salesRepEmail},${salesRepManagerEmail}&message=Enquiry%20initiated&topicName=Enquire%20about%20Order%20${orderId}%20"
        }
    ]
}