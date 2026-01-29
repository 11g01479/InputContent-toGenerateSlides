
import { GoogleGenAI, Type } from "@google/genai";

// TypeScript declaration for the PptxGenJS library loaded from a script tag
declare var PptxGenJS: any;

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
        reader.onload = () => resolve(reader.result as string); // result is data:mime/type;base64,...
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
    // Select all necessary DOM elements
    const scriptInput = document.getElementById('script-input') as HTMLTextAreaElement;
    const imageInput = document.getElementById('image-input') as HTMLInputElement;
    const imagePreview = document.getElementById('image-preview') as HTMLDivElement;
    const generateBtn = document.getElementById('generate-btn') as HTMLButtonElement;
    const presentationOutput = document.getElementById('presentation-output') as HTMLDivElement;

    // Early exit if any essential element is not found
    if (!scriptInput || !imageInput || !imagePreview || !generateBtn || !presentationOutput) {
        console.error("A required DOM element was not found. The application cannot start.");
        return;
    }
    
    // Variables to store generated content for download
    let finalSlidesContent: { title: string; content: string[]; imageIndex: number; imageGenerationPrompt?: string }[] | null = null;
    let finalSlideImagesData: ({ base64: string; dims: { width: number; height: number; }; } | null)[] | null = null;

    // Event listener for image input to show previews
    imageInput.addEventListener('change', () => {
        imagePreview.innerHTML = ''; // Clear previous previews
        const files = imageInput.files;
        if (!files) return;

        // Create and display a preview for each selected image file
        for (const file of files) {
            const reader = new FileReader();
            reader.onload = (e) => {
                const img = document.createElement('img');
                if (e.target?.result) {
                    img.src = e.target.result as string;
                }
                img.alt = `Preview: ${file.name}`;
                imagePreview.appendChild(img);
            };
            reader.readAsDataURL(file);
        }
    });

    // Event listener for the generate button
    generateBtn.addEventListener('click', async () => {
        const scriptText = scriptInput.value.trim();
        const imageFiles = imageInput.files;
        
        if (!scriptText) {
            alert('プレゼンテーション原稿を入力してください。');
            return;
        }
        
        // Reset stored data
        finalSlidesContent = null;
        finalSlideImagesData = null;

        // Set loading state on UI elements
        generateBtn.disabled = true;
        generateBtn.textContent = '生成中...';
        presentationOutput.innerHTML = '<p class="status-message">AIがスライド構成を分析しています。少々お待ちください...</p>';

        try {
            // Re-initialize AI client here to ensure it uses the current process.env.API_KEY
            const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

            const imageCount = imageFiles?.length ?? 0;
            const prompt = `あなたはプレゼンテーション作成の専門家です。以下のテキストと${imageCount > 0 ? `提供された${imageCount}枚の画像` : 'テキスト'}を解析し、最適なプレゼンテーションを構成してください。

各スライドについて、明確なタイトルと本文の箇条書きを作成してください。
本文の箇条書きは、内容に応じて番号付きリスト（例：1.、2.、3.）または通常の箇条書きにしてください。リストの先頭に絵文字やアイコンは使用しないでください。

各スライドの内容を考慮し、${imageCount > 0 ? `提供された画像の中から最も適切なものを割り当ててください。もし提供された画像の中に適切なものがない、あるいはさらに良い画像が考えられる場合は、新しい画像を生成するように指示してください。` : '各スライドに最適な画像を生成するように指示してください。'}

出力は、各オブジェクトがスライドを表すJSON配列として提供してください。
各オブジェクトには以下を含めてください。
- 'title': スライドのタイトル
- 'content': スライドの本文（箇条書きの配列）
- 'imageIndex': ${imageCount > 0 ? `提供された画像を使用する場合、その画像の0から始まるインデックス。新しい画像を生成する場合は -1 を指定してください。` : `常に -1 を指定してください。`}
- 'imageGenerationPrompt': 'imageIndex'が-1の場合にのみ、画像を生成するための詳細でクリエイティブな**英語のプロンプト**を含めてください。写実的な写真(photorealistic)や、モダンなイラスト(modern illustration)など、スタイルも指定してください。

${imageCount > 0 ? `提供されたすべての画像が必ずしも使われる必要はありません。内容に合わない場合は無理に使用せず、新しい画像を生成してください。` : ''}
---
${scriptText}
---`;
            
            const responseSchema = {
                type: Type.ARRAY,
                items: {
                  type: Type.OBJECT,
                  properties: {
                    title: {
                      type: Type.STRING,
                      description: 'スライドのタイトル',
                    },
                    content: {
                      type: Type.ARRAY,
                      items: {
                        type: Type.STRING,
                      },
                      description: 'スライドの本文コンテンツ（箇条書きの配列）。絵文字は含めないでください。',
                    },
                    imageIndex: {
                        type: Type.INTEGER,
                        description: `使用する提供画像の0ベースのインデックス。新しい画像を生成する場合は-1を指定してください。`,
                    },
                    imageGenerationPrompt: {
                        type: Type.STRING,
                        description: 'imageIndexが-1の場合に、新しい画像を生成するための詳細な英語のプロンプト。',
                    }
                  },
                  required: ['title', 'content', 'imageIndex'],
                },
            };

            const textPart = { text: prompt };
            const imageParts = imageFiles ? await Promise.all(Array.from(imageFiles).map(fileToGenerativePart)) : [];
            const contents = { parts: [textPart, ...imageParts] };

            // Call the Gemini API to generate content structure using Gemini 3 Flash
            const response = await ai.models.generateContent({
                model: 'gemini-3-flash-preview',
                contents: contents,
                config: {
                    responseMimeType: "application/json",
                    responseSchema,
                }
            });

            const jsonResponse = response.text.trim();
            const slidesContent: { title: string; content: string[]; imageIndex: number; imageGenerationPrompt?: string }[] = JSON.parse(jsonResponse);
            
            presentationOutput.innerHTML = '<p class="status-message">スライドの画像を準備しています...</p>';

            const imageBase64s = imageFiles ? await Promise.all(Array.from(imageFiles).map(fileToBase64)) : [];
            const imageDims = imageFiles ? await Promise.all(Array.from(imageFiles).map(getImageDimensions)) : [];

            const generatedImages: ({ base64: string; dims: { width: number; height: number; }; } | null)[] = [];
            
            // Generate images sequentially to avoid hitting API rate limits
            for (const [index, slideData] of slidesContent.entries()) {
                if (slideData.imageIndex >= 0 && imageBase64s[slideData.imageIndex]) {
                    // Use an existing image
                    generatedImages.push({
                        base64: imageBase64s[slideData.imageIndex],
                        dims: imageDims[slideData.imageIndex],
                    });
                } else if (slideData.imageGenerationPrompt) {
                    // Generate a new image using Gemini 2.5 Flash Image
                    presentationOutput.innerHTML = `<p class="status-message">スライド ${index + 1}/${slidesContent.length} の画像を生成中...<br><small>${slideData.imageGenerationPrompt}</small></p>`;
                    try {
                        const imageResponse = await ai.models.generateContent({
                            model: 'gemini-2.5-flash-image',
                            contents: {
                                parts: [{ text: slideData.imageGenerationPrompt }]
                            },
                            config: {
                                imageConfig: {
                                    aspectRatio: "16:9"
                                }
                            }
                        });

                        let foundImage = false;
                        if (imageResponse.candidates?.[0]?.content?.parts) {
                            for (const part of imageResponse.candidates[0].content.parts) {
                                if (part.inlineData) {
                                    const base64ImageBytes = part.inlineData.data;
                                    const imageUrl = `data:${part.inlineData.mimeType};base64,${base64ImageBytes}`;
                                    const dims = await getBase64ImageDimensions(base64ImageBytes);
                                    generatedImages.push({ base64: imageUrl, dims });
                                    foundImage = true;
                                    break;
                                }
                            }
                        }
            
                        if (!foundImage) {
                            console.warn(`No image returned in response for slide ${index + 1}`);
                            generatedImages.push(null);
                        }
            
                    } catch (genError) {
                        console.error(`Error generating image for slide ${index + 1}:`, genError);
                        let errorMessage = `スライド ${index + 1} の画像生成に失敗しました。`;
                        if (String(genError).includes('429')) {
                            errorMessage += ' APIの利用制限に達した可能性があります。';
                        }
                        presentationOutput.innerHTML += `<p style="color: orange; font-size: 0.8rem;">${errorMessage} このスライドは画像なしで作成されます。</p>`;
                        generatedImages.push(null);
                    }
                } else {
                    // No image for this slide
                    generatedImages.push(null);
                }
            }

            // Store generated data for the download button
            finalSlidesContent = slidesContent;
            finalSlideImagesData = generatedImages;

            // --- Render Slide Previews ---
            presentationOutput.innerHTML = ''; // Clear status message
            const previewHeader = document.createElement('h2');
            previewHeader.textContent = '生成されたスライドのプレビュー';
            presentationOutput.appendChild(previewHeader);

            const previewContainer = document.createElement('div');
            previewContainer.id = 'slide-previews-container';
            presentationOutput.appendChild(previewContainer);

            slidesContent.forEach((slideData, index) => {
                const slidePreview = document.createElement('div');
                slidePreview.className = 'slide-preview';

                const slideNumber = document.createElement('span');
                slideNumber.className = 'slide-number';
                slideNumber.textContent = `スライド ${index + 1}`;
                slidePreview.appendChild(slideNumber);
    
                const slideTitle = document.createElement('h3');
                slideTitle.textContent = slideData.title;
                slidePreview.appendChild(slideTitle);
    
                const slideContentPreview = document.createElement('div');
                slideContentPreview.className = 'slide-content-preview';
                slideContentPreview.innerHTML = slideData.content.map(p => `<p>${p.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</p>`).join('');
                slidePreview.appendChild(slideContentPreview);
    
                const imageData = generatedImages[index];
                if (imageData) {
                    const img = document.createElement('img');
                    img.src = imageData.base64;
                    img.alt = `Slide ${index + 1} image preview`;
                    slidePreview.appendChild(img);
                }
    
                previewContainer.appendChild(slidePreview);
            });

            // --- Add Download Button ---
            const downloadButton = document.createElement('button');
            downloadButton.id = 'download-btn';
            downloadButton.className = 'action-button';
            downloadButton.textContent = 'PowerPointをダウンロード';
            downloadButton.style.marginTop = '2rem';
            presentationOutput.appendChild(downloadButton);

        } catch (error) {
            console.error("Error generating presentation:", error);
            presentationOutput.innerHTML = '<p style="color: red;">生成中にエラーが発生しました。設定や入力内容を確認してください。</p>';
            alert('プレゼンテーションの生成中にエラーが発生しました。');
        } finally {
            // Reset UI elements from loading state
            generateBtn.disabled = false;
            generateBtn.textContent = 'プレゼンテーションを生成';
        }
    });

    // Event listener for the download button (using event delegation)
    presentationOutput.addEventListener('click', async (event) => {
        const target = event.target as HTMLElement;
        if (target.id !== 'download-btn' || !finalSlidesContent || !finalSlideImagesData) {
            return;
        }

        const downloadBtn = target as HTMLButtonElement;
        downloadBtn.disabled = true;
        downloadBtn.textContent = 'PowerPointファイルを生成中...';

        try {
            // Create a new PowerPoint presentation
            const pptx = new PptxGenJS();
            
            finalSlidesContent.forEach((slideData, slideIndex) => {
                const slide = pptx.addSlide();

                // Add title
                slide.addText(slideData.title, { 
                    x: 0.5, y: 0.25, w: 9, h: 0.75, 
                    fontSize: 28, bold: true, color: '00529B' 
                });

                const imageData = finalSlideImagesData[slideIndex];
                const textContent = slideData.content.join('\n\n');
                
                if (imageData) {
                    // --- SLIDE WITH IMAGE ---
                    const { base64: imageBase64, dims } = imageData;
                    const { width: originalWidth, height: originalHeight } = dims;
                    
                    const layoutType = slideIndex % 2 === 0 ? 'imageRight' : 'imageLeft';

                    let textOptions: any;
                    let imageArea: { x: number; y: number; w: number; h: number; };

                    if (layoutType === 'imageRight') {
                        textOptions = { x: 0.5, y: 1.2, w: 5.0, h: 4.2, fontSize: 14, valign: 'top' };
                        imageArea = { x: 6.0, y: 1.2, w: 3.5, h: 4.2 };
                    } else { // imageLeft
                        textOptions = { x: 4.5, y: 1.2, w: 5.0, h: 4.2, fontSize: 14, valign: 'top' };
                        imageArea = { x: 0.5, y: 1.2, w: 3.5, h: 4.2 };
                    }
                    
                    slide.addText(textContent, textOptions);

                    const aspectRatio = originalWidth / originalHeight;
                    const maxBoxAspectRatio = imageArea.w / imageArea.h;

                    let newWidth, newHeight;
                    if (aspectRatio > maxBoxAspectRatio) {
                        newWidth = imageArea.w;
                        newHeight = newWidth / aspectRatio;
                    } else {
                        newHeight = imageArea.h;
                        newWidth = newHeight * aspectRatio;
                    }

                    const newX = imageArea.x + (imageArea.w - newWidth) / 2;
                    const newY = imageArea.y + (imageArea.h - newHeight) / 2;

                    slide.addImage({ 
                        data: imageBase64, 
                        x: newX, y: newY, w: newWidth, h: newHeight
                    });
                } else {
                    // --- SLIDE WITHOUT IMAGE (TEXT ONLY) ---
                    slide.addText(textContent, { 
                        x: 0.5, y: 1.2, w: 9, h: 4.2, 
                        fontSize: 16, valign: 'top' 
                    });
                }
            });

            await pptx.writeFile({ fileName: 'AI-Generated-Presentation.pptx' });
            
            downloadBtn.textContent = 'ダウンロード完了！';
            downloadBtn.style.backgroundColor = '#28a745';

        } catch (error) {
            console.error("Error writing PowerPoint file:", error);
            downloadBtn.textContent = 'エラーが発生しました';
            downloadBtn.style.backgroundColor = '#dc3545';
            alert('PowerPointファイルの生成中にエラーが発生しました。');
        }
    });
});
