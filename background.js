function getPromptForOption(option) {
    const prompts = {
        'A': `請你扮演一個研究者。XXXXX。`,
        'B': `## 請你扮演一個幫助釐清問題的助手。XXXXX`,
 'C': `請根據以下XXXXX`,

'E':`請根據以下XXXXX`,
'F':`我將給你多個詞彙以及其對應的XXXXX`,
'G':`請你扮演一個寫作建議師，我將給你一段XXXX
`,
'H':`不管我傳什麼,都回我哈哈哈。
`

};
    return prompts[option] || '';
}

chrome.runtime.onMessage.addListener(function(request, sender, sendResponse) {
    if (request.action === "generateReply" || request.action === "analyzeExcel") {

        const apiKey = "sk-XXXXX"; 

        let systemPrompt = getPromptForOption(request.option);
        let fullPrompt;

        if (request.action === "generateReply") {
            fullPrompt = systemPrompt + "\n\n" + request.userText;
        } else if (request.action === "analyzeExcel") {
            fullPrompt = systemPrompt + "\n\n" + JSON.stringify(request.data);
        } else if (request.action === "generateCustomReply") {
            fullPrompt = request.systemPrompt +"\n\n" + request.userPrompt;
        }

        fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${apiKey}`
            },
            body: JSON.stringify({
                temperature: 0.7,
                model: "gpt-4o",
                messages: [{
                    role: "system",
                    content: systemPrompt
                }, {
                    role: "user",
                    content: fullPrompt
                }]
            })
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`API 請求失敗，狀態碼：${response.status}`);
            }
            return response.json();
        })
        .then(data => {
            if (data.choices && data.choices.length > 0 && data.choices[0].message) {
                sendResponse({ result: data.choices[0].message.content });
            } else {
                sendResponse({ error: "未收到有效回覆" });
            }
        })
        .catch(error => {
            sendResponse({ error: "API 調用失敗：" + error.message });
        });

        return true; 
    } else if (request.action === "calculate") {

        try {
            const labels = JSON.parse(request.labels);
            const res90 = request.res90;
            const res30 = request.res30;

            let resultText = calculateLabelInterestChange(labels, res90, res30);
            sendResponse({ result: resultText });
        } catch (e) {
            sendResponse({ error: "输入的JSON格式错误或解析失败。" });
        }
    }
});


chrome.browserAction.onClicked.addListener(function() {

    chrome.windows.create({
        url: chrome.runtime.getURL("popup.html"),
        type: "popup",
        width: 400,
        height: 600
    });
});

