import { createWorker } from 'tesseract.js';
import { Document, Packer, Paragraph, TextRun } from 'docx';
import { saveAs } from 'file-saver';

let selectedFiles = [];
let worker = null;

// Éléments de l'interface utilisateur
const progressBar = document.getElementById('progress');
const statusMessage = document.getElementById('statusMessage');
const fileNameDisplay = document.getElementById('file-name');
const convertButton = document.getElementById('convert');

// Gestion de la police
let fontFamily = 'Times New Roman';
let fontSize = 12;

// Compteurs globaux pour la nomenclature unique
let personCounter = 1;
let companyCounter = 1;

async function initializeWorker() {
    if (!worker) {
        worker = await createWorker('fra'); // Initialise avec la langue française
    }
}

async function adjustImageContrast(file, contrast = 70) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = function(event) {
            const img = new Image();
            img.onload = function() {
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                canvas.width = img.width;
                canvas.height = img.height;
                
                ctx.drawImage(img, 0, 0);
                
                const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                const data = imageData.data;
                
                const factor = (259 * (contrast + 255)) / (255 * (259 - contrast));
                
                for (let i = 0; i < data.length; i += 4) {
                    data[i] = factor * (data[i] - 128) + 128;
                    data[i + 1] = factor * (data[i + 1] - 128) + 128;
                    data[i + 2] = factor * (data[i + 2] - 128) + 128;
                }
                
                ctx.putImageData(imageData, 0, 0);
                
                canvas.toBlob((blob) => {
                    resolve(new File([blob], file.name, { type: 'image/png' }));
                }, 'image/png');
            };
            img.src = event.target.result;
        };
        reader.readAsDataURL(file);
    });
}

async function extractTextFromImage(file) {
    await initializeWorker();

    try {
        // Ajuste le contraste avant l'OCR
        const adjustedFile = await adjustImageContrast(file, 50); // Augmente le contraste de 50%

        const { data } = await worker.recognize(adjustedFile, {
            rotate: 'auto', // Active la rotation automatique pour une meilleure précision
        });

        return {
            text: smartAnonymize(data.text),
            confidence: data.confidence,
            words: data.words.map(word => ({
                ...word,
                text: smartAnonymize(word.text)
            }))
        };
    } catch (error) {
        console.error('Erreur OCR:', error);
        throw error;
    }
}

function smartAnonymize(text) {
    const nameMap = new Map();
    const companyMap = new Map();

    return text.replace(/\[([^\]]+)\]/g, (match, name) => {
        // Vérifie si c'est probablement un nom d'entreprise
        if (name.includes('Banque') || name.includes('Société') || name.includes('SARL') || name.includes('SA')) {
            if (!companyMap.has(name)) {
                companyMap.set(name, `[ENTREPRISE-${companyCounter++}]`);
            }
            return companyMap.get(name);
        } else {
            // Suppose que c'est un nom de personne
            if (!nameMap.has(name)) {
                const initials = name.split(' ').map(part => part[0]).join('');
                nameMap.set(name, `[PERSONNE-${initials}-${personCounter++}]`);
            }
            return nameMap.get(name);
        }
    });
}

async function processImages(files) {
    const results = [];
    const totalFiles = files.length;

    for (let i = 0; i < totalFiles; i++) {
        const file = files[i];
        statusMessage.textContent = `Traitement de ${file.name} (${i + 1} sur ${totalFiles})...`;
        const result = await extractTextFromImage(file);
        results.push({ filename: file.name, ...result });

        const progressPercentage = ((i + 1) / totalFiles) * 100;
        progressBar.style.width = `${progressPercentage}%`;
    }

    return results;
}

async function createWordFile(extractedTexts) {
    const doc = new Document({
        sections: [],
        defaultStyle: {
            font: fontFamily,
            size: fontSize * 2 // docx utilise des demi-points pour la taille de police
        }
    });

    extractedTexts.forEach(({ filename, text, confidence }) => {
        doc.addSection({
            children: [
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `Image : ${filename}`,
                            bold: true,
                        }),
                    ],
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: `Confiance : ${confidence.toFixed(2)}%`,
                            italic: true,
                            bold: true,
                        }),
                    ],
                }),
                ...text.split('\n').map(line => 
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: line,
                            }),
                        ],
                    })
                ),
            ],
        });
    });

    const buffer = await Packer.toBlob(doc);
    saveAs(buffer, 'texte_extrait.docx');
}

// Gestion de la sélection de fichiers
document.getElementById('upload').addEventListener('change', (event) => {
    selectedFiles = event.target.files;
    progressBar.style.width = '0%';
    statusMessage.textContent = '';
    
    if (selectedFiles.length > 0) {
        fileNameDisplay.textContent = selectedFiles.length === 1 
            ? selectedFiles[0].name 
            : `${selectedFiles.length} fichiers sélectionnés`;
        convertButton.disabled = false;
    } else {
        fileNameDisplay.textContent = 'Aucun fichier choisi';
        convertButton.disabled = true;
    }
});

// Gestion du changement de police
document.getElementById('fontFamily').addEventListener('change', (event) => {
    fontFamily = event.target.value;
});

// Gestion du changement de taille de police
document.getElementById('fontSize').addEventListener('change', (event) => {
    fontSize = parseInt(event.target.value);
});

// Gestion de la conversion lors du clic sur le bouton
convertButton.addEventListener('click', async () => {
    if (selectedFiles.length === 0) {
        alert('Veuillez sélectionner une ou plusieurs images.');
        return;
    }

    convertButton.disabled = true;
    statusMessage.textContent = 'Traitement des images en cours...';
    progressBar.style.width = '0%';

    try {
        const extractedTexts = await processImages(selectedFiles);
        
        statusMessage.textContent = 'Création du document Word...';
        await createWordFile(extractedTexts);
        
        statusMessage.textContent = 'Conversion terminée ! Votre fichier Word est prêt.';
    } catch (error) {
        statusMessage.textContent = 'Une erreur est survenue pendant le traitement.';
        console.error('Erreur de traitement:', error);
    } finally {
        convertButton.disabled = false;
    }
});

// Fonction de nettoyage pour terminer le worker lorsque l'application se ferme
async function cleanup() {
    if (worker) {
        await worker.terminate();
        worker = null;
    }
}

// Appel de la fonction de nettoyage lorsque la fenêtre est sur le point de se décharger
window.addEventListener('beforeunload', cleanup);

// État initial du bouton
convertButton.disabled = true;