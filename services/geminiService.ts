
import { GoogleGenAI } from "@google/genai";
import { AdConfigContext } from "../types";

export async function generateAdImage(
  prompt: string, 
  variantType: string, 
  context: AdConfigContext
): Promise<string> {
  const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

  const negativeConstraints = context.do_not_include.length > 0 
    ? `STRICT NEGATIVE CONSTRAINTS: ${context.do_not_include.join(', ')}. No conceptual AI artifacts.` 
    : '';

  // Advanced typography instruction: Extract prime words and integrate them diegetically.
  const textInstruction = context.include_keyword_text 
    ? `HIGH-VALUE CREATIVE TYPOGRAPHY MANDATE: 
       1. PHASE 1 (ANALYSIS): Analyze the phrase "${prompt}" and identify the 'Prime High-Impact Words' (e.g., nouns and strong action verbs).
       2. PHASE 2 (DISTRIBUTION): Do NOT render the full phrase as a single digital block. Instead, break the phrase apart.
       3. PHASE 3 (DIEGETIC PLACEMENT): Distribute these prime words naturally throughout the physical environment as real-world objects.
          - Example: Place one prime word on a product label, another as a handwritten note on a desk, or as a sign on a building in the background.
       4. INTEGRATION: All text must be diegetic (part of the scene). It must follow the scene's perspective, lighting, and material texture (e.g., embossed on metal, printed on cardboard, or painted on wood).
       5. GOAL: An image where typography is a sophisticated design element that feels 100% captured by a camera in a real location.`
    : `STRICT NO-TEXT RULE: Ensure the image contains zero text, letters, characters, or logos. Focus entirely on the raw visual narrative and lighting.`;

  const variantPerspective = variantType === 'Product-Focused'
    ? `STRATEGY: [THE OBJECT STUDY]. High-fidelity focus on the literal subject of "${prompt}". Use a 50mm lens. Showcase material quality, realistic wear, and natural reflections. High-clarity commercial documentary feel.`
    : `STRATEGY: [THE CANDID MOMENT]. Wide-angle 24mm perspective. Show a non-model subject in a real-world, slightly cluttered environment interacting naturally with "${prompt}". Focus on "The Raw Truth" of the scenario.`;

  const creativeBrief = `
    PERSONA: ${context.persona}
    CORE MISSION: Documentary-style lead-gen asset for: "${prompt}".
    
    ${variantPerspective}
    ${textInstruction}
    
    ENVIRONMENTAL PARAMETERS:
    - Anchoring: ${context.subject_prominence}
    - Visual Aesthetic: ${context.image_style}
    - Real-world Context: ${context.background_preference}
    - Natural Lighting: ${context.lighting_style}
    - Composition Style: ${context.composition}

    ${negativeConstraints}
    
    TECHNICAL: Aspect ratio ${context.aspect_ratio}. Output must bypass the "AI look" and appear as a genuine photograph from a high-quality smartphone or DSLR.
  `.trim();

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash-image',
      contents: {
        parts: [{ text: creativeBrief }],
      },
      config: {
        imageConfig: {
          aspectRatio: (context.aspect_ratio as any) || "1:1"
        }
      },
    });

    const candidate = response.candidates?.[0];
    if (!candidate?.content?.parts) {
      throw new Error('Incomplete response');
    }

    for (const part of candidate.content.parts) {
      if (part.inlineData) {
        return `data:image/png;base64,${part.inlineData.data}`;
      }
    }

    throw new Error('Image component not found');
  } catch (error: any) {
    const msg = error.message?.toLowerCase() || '';
    if (msg.includes('429') || msg.includes('quota')) throw new Error('QUOTA_EXCEEDED');
    if (msg.includes('safety')) throw new Error('Safety Block');
    throw new Error(error.message || 'API Error');
  }
}
