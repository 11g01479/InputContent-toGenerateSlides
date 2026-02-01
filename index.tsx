
import { GoogleGenAI, Type } from "@google/genai";

// TypeScript declaration for the PptxGenJS library loaded from a script tag
declare var PptxGenJS: any;

// Constants for quota management
const DAILY_QUOTA_LIMIT = 100;
const STORAGE_KEY_COUNT = 'ai_presentation_gen_count';
const STORAGE_KEY_DATE = 'ai_presentation_gen_date';

// Utility function to wait for a specified time
const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

// Quota Helper Functions
function getRemainingQuota(): number {
    const today = new Date().toLocaleDateString();
    const storedDate = localStorage.getItem(STORAGE_KEY_DATE);
    
    if (storedDate !== today) {
        localStorage.setItem(STORAGE_KEY_DATE, today);
        localStorage.setItem(STORAGE_KEY_COUNT, '0');
        return DAILY_QUOTA_LIMIT;
    }
    
    const count = parseInt(localStorage.getItem(STORAGE_KEY_COUNT) || '0', 10);
    return Math.max(0, DAILY_QUOTA_LIMIT - count);
}

function incrementUsage() {
    const count = parseInt(localStorage.getItem(STORAGE_KEY_COUNT) || '0', 10);
    localStorage.setItem(STORAGE_KEY_COUNT, (count + 1).toString());
    updateQuotaDisplay();
}

function updateQuotaDisplay() {
    const remaining = getRemainingQuota();
    const displayElement = document.getElementById('remaining-quota-display');
    if (displayElement) {
        displayElement.textContent = `本日あと ${remaining} 回利用可能`;
        if (remaining <= 0) {
            displayElement.style.color = 'var(--error-color)';
            const genBtn = document.getElementById('generate-btn') as HTMLButtonElement;
            if (genBtn) genBtn.disabled = true;
        }
    }
}

// Helper function to convert a File object to a GoogleGenerativeAI.Part object
async function fileToGenerativePart(file: File) {
    const base64EncodedDataPromise = new Promise<string>((resolve) => {
        const reader = new FileReader();
        reader.onloadend = () => {
            const base64Data = (reader.result as string).split(',')[1];
            resolve(base64Data);
        };
        reader.readAsDataURL(file);
    });

    return {
        inlineData: {
            data: await base64EncodedDataPromise,
            mimeType: file.type,
        },
    };
}

// Helper function to convert a File to a Base64 string for PptxGenJS
function fileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = error => reject(error);
        reader.readAsDataURL(file);
    });
}

// Helper function to get image dimensions from a File object
function getImageDimensions(file: File): Promise<{ width: number; height: number }> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            if (!e.target?.result) {
                return reject(new Error("Could not read file for dimensions."));
            }
            const img = new Image();
            img.onload = () => {
                resolve({ width: img.naturalWidth, height: img.naturalHeight });
            };
            img.onerror = (err) => reject(err);
            img.src = e.target.result as string;
        };
        reader.onerror = (err) => reject(err);
        reader.readAsDataURL(file);
    });
}

// Helper function to get image dimensions from a base64 string
function getBase64ImageDimensions(base64: string): Promise<{ width: number; height: number }> {
    return new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => {
            resolve({ width: img.naturalWidth, height: img.naturalHeight });
        };
        img.onerror = (err) => reject(err);
        img.src = `data:image/png;base64,${base64}`;
    });
}


// Ensure the DOM is fully loaded before running the script
document.addEventListener('DOMContentLoaded', () => {
    const scriptInput = document.getElementById('script-input') as HTMLTextAreaElement;
    const imageInput = document.getElementById('image-input') as HTMLInputElement;
    const imagePreview = document.getElementById('image-preview') as HTMLDivElement;
    const generateBtn = document.getElementById('generate-btn') as HTMLButtonElement;
    const presentationOutput = document.getElementById('presentation-output') as HTMLDivElement;

    // Initial quota display
    updateQuotaDisplay();

    if (!scriptInput || !imageInput || !imagePreview || !generateBtn || !presentationOutput) {
        console.error("Required DOM elements missing.");
        return;
    }
    
    let finalSlidesContent: { title: string; content: string[]; imageIndex: number; imageGenerationPrompt?: string }[] | null = null;
    let finalSlideImagesData: ({ base64: string; dims: { width: number; height: number; }; } | null)[] | null = null;

    imageInput.addEventListener('change', () => {
        imagePreview.innerHTML = '';
        const files = imageInput.files;
        if (!files) return;

        for (const file of files) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const img = document.createElement('img');
                if (e.target?.result) img.src = e.target.result as string;
                img.alt = `Preview: ${file.name}`;
                imagePreview.appendChild(img);
            };
            reader.readAsDataURL(file);
        }
    });

    generateBtn.addEventListener('click', async () => {
        if (getRemainingQuota() <= 0) {
            alert('本日の利用上限に達しました。明日またお試しください。');
            return;
        }

        const scriptText = scriptInput.value.trim();
        const imageFiles = imageInput.files;
        
        if (!scriptText) {
            alert('プレゼンテーション原稿を入力してください。');
            return;
        }
        
        finalSlidesContent = null;
        finalSlideImagesData = null;
        generateBtn.disabled = true;
        generateBtn.textContent = '生成中...';
        presentationOutput.innerHTML = '<p class="status-message">AIがスライド構成を分析しています...</p>';

        try {
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
            const imageCount = imageFiles?.length ?? 0;
            
            const prompt = `あなたはプレゼンテーション作成の専門家です。以下のテキストを解析し、各スライドに必ず1枚の画像が含まれる魅力的なプレゼンテーション（スライド構成）を作成してください。

**最優先の制約事項:**
1. すべてのスライドに必ず画像が含まれている必要があります。
2. 提供された画像（${imageCount}枚）がある場合、内容に合うものがあればそれを使用してください。
3. 提供された画像がない場合、または提供された画像が内容に合わない場合は、必ず 'imageIndex' を -1 に設定し、そのスライドに最適な画像を生成するための詳細な英語プロンプトを 'imageGenerationPrompt' に作成してください。

各スライドについて：
- 'title': スライドのタイトル
- 'content': 箇条書きの本文（配列）
- 'imageIndex': 提供画像を使う場合は 0 からのインデックス。新規生成なら -1。
- 'imageGenerationPrompt': 'imageIndex' が -1 の場合に必須。写真のような高品質な英語プロンプト。

---
${scriptText}
---`;
            
            const responseSchema = {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    title: { type: Type.STRING },
                    content: { type: Type.ARRAY, items: { type: Type.STRING } },
                    imageIndex: { type: Type.INTEGER },
                    imageGenerationPrompt: { type: Type.STRING, description: "imageIndexが-1の場合に必須の詳細な英語プロンプト" }
                  },
                  required: ['title', 'content', 'imageIndex', 'imageGenerationPrompt'],
                },
            };

            const textPart = { text: prompt };
            const imageParts = imageFiles ? await Promise.all(Array.from(imageFiles).map(fileToGenerativePart)) : [];
            
            const response = await ai.models.generateContent({
                model: 'gemini-3-flash-preview',
                contents: { parts: [textPart, ...imageParts] },
                config: {
                    responseMimeType: "application/json",
                    responseSchema,
                }
            });

            const slidesContent: { title: string; content: string[]; imageIndex: number; imageGenerationPrompt?: string }[] = JSON.parse(response.text.trim());
            
            presentationOutput.innerHTML = '<p class="status-message">スライドの画像を準備しています...</p>';

            const imageBase64s = imageFiles ? await Promise.all(Array.from(imageFiles).map(fileToBase64)) : [];
            const imageDims = imageFiles ? await Promise.all(Array.from(imageFiles).map(getImageDimensions)) : [];
            const generatedImages: ({ base64: string; dims: { width: number; height: number; }; } | null)[] = [];
            
            for (const [index, slideData] of slidesContent.entries()) {
                if (index > 0) await sleep(2000);

                if (slideData.imageIndex >= 0 && imageBase64s[slideData.imageIndex]) {
                    generatedImages.push({
                        base64: imageBase64s[slideData.imageIndex],
                        dims: imageDims[slideData.imageIndex],
                    });
                } else {
                    const genPrompt = slideData.imageGenerationPrompt || `A high quality, professional presentation slide image for "${slideData.title}", modern style.`;
                    presentationOutput.innerHTML = `<p class="status-message">スライド ${index + 1}/${slidesContent.length} の画像を生成中...<br><small>API制限を考慮しつつ1枚ずつ作成しています</small></p>`;
                    
                    try {
                        const imageResponse = await ai.models.generateContent({
                            model: 'gemini-2.5-flash-image',
                            contents: { parts: [{ text: genPrompt }] },
                            config: { imageConfig: { aspectRatio: "16:9" } }
                        });

                        let found = false;
                        if (imageResponse.candidates?.[0]?.content?.parts) {
                            for (const part of imageResponse.candidates[0].content.parts) {
                                if (part.inlineData) {
                                    const dims = await getBase64ImageDimensions(part.inlineData.data);
                                    generatedImages.push({ base64: `data:${part.inlineData.mimeType};base64,${part.inlineData.data}`, dims });
                                    found = true;
                                    break;
                                }
                            }
                        }
                        if (!found) generatedImages.push(null);
                    } catch (err) {
                        console.error(err);
                        generatedImages.push(null);
                    }
                }
            }

            finalSlidesContent = slidesContent;
            finalSlideImagesData = generatedImages;

            presentationOutput.innerHTML = '<h2>プレビュー</h2><div id="slide-previews-container"></div>';
            const container = document.getElementById('slide-previews-container')!;

            slidesContent.forEach((slideData, index) => {
                const div = document.createElement('div');
                div.className = 'slide-preview';
                div.innerHTML = `
                    <span class="slide-number">スライド ${index + 1}</span>
                    <h3>${slideData.title}</h3>
                    <div class="slide-content-preview">${slideData.content.map(p => `<p>${p}</p>`).join('')}</div>
                `;
                if (generatedImages[index]) {
                    const img = document.createElement('img');
                    img.src = generatedImages[index]!.base64;
                    div.appendChild(img);
                } else {
                    div.innerHTML += `<p style="color: #999; font-size: 0.8rem; border: 1px dashed #ccc; padding: 1rem; text-align: center;">画像なし（生成エラー）</p>`;
                }
                container.appendChild(div);
            });

            const dlBtn = document.createElement('button');
            dlBtn.id = 'download-btn';
            dlBtn.className = 'action-button';
            dlBtn.textContent = 'PowerPointをダウンロード';
            dlBtn.style.marginTop = '2rem';
            presentationOutput.appendChild(dlBtn);

            // Successfully finished generation, increment usage
            incrementUsage();

        } catch (error) {
            console.error(error);
            let msg = '生成中にエラーが発生しました。';
            if (String(error).includes('429')) msg = 'APIの利用制限（1分あたりの上限）に達しました。少し時間を置いてから再度お試しください。';
            presentationOutput.innerHTML = `<p style="color: red;">${msg}</p>`;
        } finally {
            generateBtn.disabled = (getRemainingQuota() <= 0);
            generateBtn.textContent = 'プレゼンテーションを生成';
        }
    });

    presentationOutput.addEventListener('click', async (event) => {
        const target = event.target as HTMLElement;
        if (target.id !== 'download-btn' || !finalSlidesContent || !finalSlideImagesData) return;

        const btn = target as HTMLButtonElement;
        btn.disabled = true;
        btn.textContent = 'ファイルを準備中...';

        try {
            const pptx = new PptxGenJS();
            finalSlidesContent.forEach((slideData, i) => {
                const slide = pptx.addSlide();
                slide.addText(slideData.title, { x: 0.5, y: 0.25, w: 9, h: 0.75, fontSize: 24, bold: true, color: '00529B' });
                const img = finalSlideImagesData![i];
                if (img) {
                    const isEven = i % 2 === 0;
                    if (isEven) {
                        slide.addText(slideData.content.join('\n\n'), { x: 0.5, y: 1.2, w: 4.5, h: 4, fontSize: 14, valign: 'top' });
                        slide.addImage({ data: img.base64, x: 5.2, y: 1.2, w: 4.3, h: 3.5 });
                    } else {
                        slide.addImage({ data: img.base64, x: 0.5, y: 1.2, w: 4.3, h: 3.5 });
                        slide.addText(slideData.content.join('\n\n'), { x: 5.0, y: 1.2, w: 4.5, h: 4, fontSize: 14, valign: 'top' });
                    }
                } else {
                    slide.addText(slideData.content.join('\n\n'), { x: 0.5, y: 1.2, w: 9, h: 4, fontSize: 16, valign: 'top' });
                }
            });
            await pptx.writeFile({ fileName: 'AI_Presentation.pptx' });
            btn.textContent = 'ダウンロード完了';
        } catch (err) {
            console.error(err);
            btn.textContent = 'エラー発生';
        } finally {
            btn.disabled = false;
        }
    });
});
