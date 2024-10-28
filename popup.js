document.addEventListener('DOMContentLoaded', function() {
    // 加载 XLSX 库
    var script = document.createElement('script');
    script.src = chrome.runtime.getURL('xlsx.full.min.js');
    document.head.appendChild(script);

    // 甲
    document.getElementById('generateButtonA').addEventListener('click', function() {
        var selectedOption = document.getElementById('optionSelectA').value;
        var userText = document.getElementById('userTextInputA').value;

        // 顯示載入中消息
        document.getElementById('loadingA').style.display = 'block';

        chrome.runtime.sendMessage({
            action: "generateReply",
            option: selectedOption,
            userText: userText
        }, function(response) {
            // 隱藏載入中消息
            document.getElementById('loadingA').style.display = 'none';

            if (chrome.runtime.lastError || !response) {
                document.getElementById('responseA').textContent = "错误：未能获取响应。";
            } else if (response.error) {
                document.getElementById('responseA').textContent = "错误：" + response.error;
            } else {
                const formattedResponse = response.result.replace(/\n/g, '<br>');
                document.getElementById('responseA').innerHTML = formattedResponse;
            }
        });
    });

    // 乙
    document.getElementById('generateButtonB').addEventListener('click', function() {
        var selectedOption = document.getElementById('optionSelectB').value;
        var userText = document.getElementById('userTextInputB').value;

        // 顯示載入中消息
        document.getElementById('loadingB').style.display = 'block';

        chrome.runtime.sendMessage({
            action: "generateReply",
            option: selectedOption,
            userText: userText
        }, function(response) {
            // 隱藏載入中消息
            document.getElementById('loadingB').style.display = 'none';

            if (chrome.runtime.lastError || !response) {
                document.getElementById('responseB').textContent = "错误：未能获取响应。";
            } else if (response.error) {
                document.getElementById('responseB').textContent = "错误：" + response.error;
            } else {
                const formattedResponse = response.result.replace(/\n/g, '<br>');
                document.getElementById('responseB').innerHTML = formattedResponse;
            }
        });
    });

    // 丙
    document.getElementById('generateButtonC').addEventListener('click', function() {
        var selectedOption = document.getElementById('optionSelectC').value;
        var userText = document.getElementById('userTextInputC').value;

        // 顯示載入中消息
        document.getElementById('loadingC').style.display = 'block';

        chrome.runtime.sendMessage({
            action: "generateReply",
            option: selectedOption,
            userText: userText
        }, function(response) {
            // 隱藏載入中消息
            document.getElementById('loadingC').style.display = 'none';

            if (chrome.runtime.lastError || !response) {
                document.getElementById('responseC').textContent = "错误：未能获取响应。";
            } else if (response.error) {
                document.getElementById('responseC').textContent = "错误：" + response.error;
            } else {
                const formattedResponse = response.result.replace(/\n/g, '<br>');
                document.getElementById('responseC').innerHTML = formattedResponse;
            }
        });
    });

    // 自訂 prompt
    document.getElementById('generateCustomReplyBtn').addEventListener('click', function() {
        var systemPrompt = document.getElementById('systemPromptText').value;
        var userPrompt = document.getElementById('userPromptText').value;

        // 顯示載入中消息
        document.getElementById('loadingCustom').style.display = 'block';

        chrome.runtime.sendMessage({
            action: "generateCustomReply",
            systemPrompt: systemPrompt,
            userPrompt: userPrompt
        }, function(response) {
            // 隱藏載入中消息
            document.getElementById('loadingCustom').style.display = 'none';

            if (chrome.runtime.lastError || !response) {
                document.getElementById('customResponse').textContent = "错误：未能获取响应。";
            } else if (response.error) {
                document.getElementById('customResponse').textContent = "错误：" + response.error;
            } else {
                const formattedResponse = response.result.replace(/\n/g, '<br>');
                document.getElementById('customResponse').innerHTML = formattedResponse;
            }
        });
    });

    // Excel 檔案上傳和分析
    document.getElementById('analyzeButton').addEventListener('click', function() {
        var fileInput = document.getElementById('uploadExcel');
        var selectedOption = document.getElementById('optionSelectExcel').value;
        if (fileInput.files.length === 0) {
            alert('請先上傳Excel檔案');
            return;
        }

        var file = fileInput.files[0];
        var reader = new FileReader();

        reader.onload = function(e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, { type: 'array' });

            var firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            var jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            document.getElementById('loadingExcel').style.display = 'block';

            // 发送到 GPT API 进行分析
            chrome.runtime.sendMessage({
                action: "analyzeExcel",
                option: selectedOption,
                data: jsonData
            }, function(response) {
                document.getElementById('loadingExcel').style.display = 'none';

                if (chrome.runtime.lastError || !response) {
                    document.getElementById('excelContent').textContent = "错误：未能获取响应。";
                } else if (response.error) {
                    document.getElementById('excelContent').textContent = "错误：" + response.error;
                } else {
                    const formattedResponse = response.result.replace(/\n/g, '<br>');
                    document.getElementById('excelContent').innerHTML = formattedResponse;
                }
            });
        };

        reader.readAsArrayBuffer(file);
    });

    // 新增工作區塊可以剪貼文字
    document.getElementById("copyTextBtn").addEventListener("click", function() {
        var textArea = document.getElementById("workAreaText");
        textArea.select();
        document.execCommand("copy");
        alert("文字已複製到剪貼簿");
    });

    document.getElementById("downloadTextBtn").addEventListener("click", function() {
        var text = document.getElementById("workAreaText").value;
        var blob = new Blob([text], { type: "text/plain;charset=utf-8" });
        var url = URL.createObjectURL(blob);
        var a = document.createElement("a");
        a.href = url;
        a.download = "workAreaText.txt";
        document.body.appendChild(a);
        a.click();
        setTimeout(function() {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 0);
    });

    // 計算關注度變化
    document.getElementById("calculate-btn").addEventListener("click", function() {
        const labelInput = document.getElementById('label_input').value;
        const res90Input = document.getElementById("res90-input").value;
        const res30Input = document.getElementById("res30-input").value;

        try {
            const labels = JSON.parse(labelInput);
            const res90 = res90Input.split(',').map(Number);
            const res30 = res30Input.split(',').map(Number);

            let resultText = calculateLabelInterestChange(labels, res90, res30);
            document.getElementById("result").innerText = resultText;
        } catch (e) {
            document.getElementById("result").innerText = "输入的JSON格式错误或解析失败。";
        }
    });

    function calculateLabelInterestChange(labels, res90, res30) {
        let result = "";

        for (let i = 0; i < 50; i++) {
            try {
                let label = labels[`label${i + 1}`];
                if (!label) break;

                let labelName = label['Name'];
                let change = res30[i] - res90[i];

                if (change > 0) {
                    result += `${labelName} 面向關注度 ${res30[i]}%（上升 ${Math.abs(change)}%）。\n`;
                } else if (change < 0) {
                    result += `${labelName} 面向關注度 ${res30[i]}%（下降 ${Math.abs(change)}%）。\n`;
                } else {
                    result += `${labelName} 面向關注度 ${res30[i]}%（不變）。\n`;
                }
            } catch (e) {
                if (e instanceof TypeError || e instanceof ReferenceError) {
                    continue;
                } else {
                    throw e;
                }
            }
        }

        return result;
    }
});
