/**
 * Simple Gemini API Client for Office Add-in
 */

/**
 * Gemini API Client Module
 */
export async function callGemini(apiKey: string, promptText: string): Promise<string> {
    if (!apiKey) throw new Error("API Key is missing. Please enter it in the taskpane.");
    
    // Using gemini-1.5-flash as the standard reliable model
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    try {
        const response = await fetch(url, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [{
                    parts: [{ text: promptText }]
                }],
                generationConfig: {
                    temperature: 0.7,
                    topP: 0.8,
                    topK: 40,
                }
            })
        });

        if (!response.ok) {
            const errorBody = await response.json();
            const errorMsg = errorBody.error?.message || "Lỗi không xác định từ API";
            throw new Error(`Gemini API Error: ${errorMsg}`);
        }

        const data = await response.json();
        const result = data.candidates?.[0]?.content?.parts?.[0]?.text;
        
        if (!result) throw new Error("API không trả về nội dung.");
        
        return result;
    } catch (error) {
        console.error("Gemini call failed:", error);
        if (error instanceof TypeError && error.message.includes("fetch")) {
            throw new Error("Không thể kết nối đến máy chủ. Vui lòng kiểm tra kết nối mạng.");
        }
        throw error;
    }
}

/**
 * Specialized Prompt Helpers
 */
export async function translateText(text: string, apiKey: string): Promise<string> {
    const prompt = `Hãy dịch đoạn văn bản sau sang tiếng Việt chuyên ngành IT/Kỹ thuật: ${text}`;
    return callGemini(apiKey, prompt);
}

export async function analyzeCode(code: string, apiKey: string): Promise<string> {
    const prompt = `Phân tích mã nguồn sau, tìm lỗi và giải thích cách sửa chi tiết: \n\n${code}`;
    return callGemini(apiKey, prompt);
}

export async function getMermaidCode(description: string, apiKey: string): Promise<string> {
    const prompt = `Hãy tạo mã MERMAID JS (Flowchart hoặc Sequence diagram) dựa trên mô tả sau. 
CHỈ TRẢ VỀ MÃ CODE MERMAID THUẦN TÚY, không kèm markdown code block, không giải thích.
Mô tả: ${description}`;
    return callGemini(apiKey, prompt);
}
