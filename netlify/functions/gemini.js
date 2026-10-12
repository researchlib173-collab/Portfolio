// Netlify Serverless Function for Gemini AI Digital Twin
// Runs securely on Netlify backend and reads process.env.GEMINI_API_KEY

exports.handler = async (event, context) => {
    // Only accept POST requests
    if (event.httpMethod !== "POST") {
        return {
            statusCode: 405,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ error: "Method Not Allowed" })
        };
    }

    try {
        const apiKey = process.env.GEMINI_API_KEY;
        if (!apiKey) {
            return {
                statusCode: 500,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ error: "GEMINI_API_KEY is not configured in Netlify Environment Variables" })
            };
        }

        const body = JSON.parse(event.body || "{}");
        const userPrompt = body.prompt || "";
        const sectionTitle = body.sectionTitle || "About";
        const sectionSummary = body.sectionSummary || "";

        const systemContext = `You are the digital twin of Balaji, a gentle, young, intelligent B.Tech CSE (IoT) student at VNR VJIET Hyderabad.
You sit in the bottom-left corner of his portfolio website as an interactive pixel-art avatar (switching between leaning against the wall with arms crossed, playing acoustic guitar, coding on a beanbag, and cozying up in a blanket burrito).
Personality: Gentle, warm, down-to-earth, witty, friendly even to rude/inappropriate questions (deflect them gently with humor).
You can speak English and casual Telugu-English (Tenglish, e.g., "Em ledu bro", "Avunu", "Chala cool", "Enti visheshalu").
Keep responses concise (2 to 4 sentences maximum) because they render inside a compact pixel comic bubble. Never use emojis.

REAL-TIME PAGE AWARENESS:
The visitor is currently viewing the "${sectionTitle}" section of the portfolio!
Section Summary: ${sectionSummary}
If the visitor asks about "this page", "here", "this section", "what am I looking at?", or refers to what is on the screen, give a curated, insightful answer directly explaining this active section!`;

        const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${encodeURIComponent(apiKey)}`;

        const response = await fetch(endpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                contents: [{
                    role: "user",
                    parts: [{ text: `${systemContext}\n\nVisitor says: ${userPrompt}` }]
                }]
            })
        });

        if (!response.ok) {
            const errText = await response.text();
            return {
                statusCode: response.status,
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ error: "Gemini API error", details: errText })
            };
        }

        const data = await response.json();
        const reply = data?.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || null;

        return {
            statusCode: 200,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ reply })
        };
    } catch (err) {
        return {
            statusCode: 500,
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ error: err.message })
        };
    }
};

