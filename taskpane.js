/*
 * RegGenAI Demo Add-in
 * AI-powered regulatory document generation
 */

// Azure Function URL - Replace with your actual deployed function URL
const AZURE_FUNCTION_URL = "https://reggenai-app-d6ehd6dddfhycch9.eastus2-01.azurewebsites.net";

// Content Control tags that should exist in the Word document
const CONTENT_CONTROLS = {
    intro: "cc_intro",
    objectives: "cc_objectives", 
    methodology: "cc_methodology"
};

// Pre-written prompts for each section
const PROMPTS = {
    intro: `You are a regulatory writer specializing in UK CTA (Clinical Trial Application) documents. Write a compelling two-paragraph introduction for a Phase 2 study of the drug 'Crinetide' for Congenital Adrenal Hyperplasia (CAH). 

The introduction should:
- Establish the medical need and burden of CAH
- Introduce Crinetide as a potential therapeutic solution
- Set up the rationale for this Phase 2 study
- Be written in formal, regulatory language suitable for UK authorities
- Be approximately 150-200 words total

Focus on the clinical and regulatory context while maintaining scientific accuracy.`,

    objectives: `You are a regulatory writer creating study objectives for a UK CTA Phase 2 study of Crinetide for Congenital Adrenal Hyperplasia. Write clear, specific primary and secondary objectives.

Primary Objectives should focus on:
- Efficacy endpoints (e.g., hormone levels, clinical symptoms)
- Safety and tolerability measures
- Dose-response relationships

Secondary Objectives should include:
- Biomarker analysis
- Quality of life measures
- Pharmacokinetic parameters

Write 3-4 primary objectives and 4-5 secondary objectives, each as a single, clear sentence. Use regulatory language appropriate for UK CTA submissions.`,

    methodology: `You are a regulatory writer drafting the methodology section for a UK CTA Phase 2 study of Crinetide for Congenital Adrenal Hyperplasia. Write a comprehensive methodology overview covering:

Study Design:
- Phase 2, randomized, double-blind, placebo-controlled study
- Multi-center design across UK sites
- Adaptive design elements for dose optimization

Patient Population:
- Adults (18-65 years) with confirmed CAH
- Specific inclusion/exclusion criteria
- Target enrollment numbers

Treatment Regimen:
- Crinetide administration details
- Placebo comparator
- Duration of treatment and follow-up

Statistical Considerations:
- Sample size justification
- Primary and secondary endpoints
- Statistical analysis plan

Write in clear, regulatory language suitable for UK authorities. Focus on the key methodological elements that would be required for CTA approval.`
};

// Initialize the add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("RegGenAI Demo Add-in loaded successfully");
        initializeAddIn();
    }
});

function initializeAddIn() {
    // Add event listeners to buttons
    document.getElementById("btn-intro").addEventListener("click", () => generateContent("intro"));
    document.getElementById("btn-objectives").addEventListener("click", () => generateContent("objectives"));
    document.getElementById("btn-methodology").addEventListener("click", () => generateContent("methodology"));
}

async function generateContent(section) {
    try {
        // Show loading state
        showStatus("Generating content...");
        disableButtons(true);

        // Get the prompt for this section
        const prompt = PROMPTS[section];
        const contentControlTag = CONTENT_CONTROLS[section];

        // For demo purposes, we'll use a mock response
        // In production, this would call your Azure Function
        const generatedText = await callAzureFunction(prompt, section);
        
        // Insert the text into the document
        await insertTextIntoDocument(generatedText, contentControlTag);
        
        // Show success state
        showStatus("Content generated successfully!", "success");
        
        // Hide status after 2 seconds
        setTimeout(() => {
            hideStatus();
            disableButtons(false);
        }, 2000);

    } catch (error) {
        console.error("Error generating content:", error);
        showStatus("Error generating content. Please try again.", "error");
        disableButtons(false);
    }
}

async function callAzureFunction(prompt, section) {
    // For the demo, we'll return mock responses
    // In production, this would make an actual HTTP call to your Azure Function
    
    const mockResponses = {
        intro: `Congenital Adrenal Hyperplasia (CAH) represents a significant unmet medical need, affecting approximately 1 in 15,000 individuals worldwide. This autosomal recessive disorder results from defects in adrenal steroidogenesis, leading to impaired cortisol production and subsequent overproduction of adrenal androgens. The clinical manifestations of CAH are severe and life-threatening, including salt-wasting crises, virilization, and impaired growth and development. Current standard of care with glucocorticoid replacement therapy, while life-saving, is associated with significant morbidity including growth suppression, obesity, and metabolic complications.

Crinetide, a novel synthetic ACTH analogue, represents a promising therapeutic approach for the treatment of CAH. By selectively stimulating cortisol production while minimizing androgen excess, Crinetide has the potential to address the fundamental pathophysiology of CAH while avoiding the adverse effects associated with supraphysiological glucocorticoid dosing. This Phase 2 study aims to evaluate the safety, tolerability, and preliminary efficacy of Crinetide in adult patients with CAH, with the goal of establishing proof-of-concept for this innovative therapeutic strategy.`,

        objectives: `Primary Objectives:
1. To evaluate the safety and tolerability of Crinetide administered subcutaneously in adult patients with Congenital Adrenal Hyperplasia over a 12-week treatment period.
2. To assess the efficacy of Crinetide in normalizing morning cortisol levels compared to placebo in patients with CAH.
3. To determine the optimal dose of Crinetide for achieving physiological cortisol levels while minimizing adverse events.
4. To evaluate the effect of Crinetide on adrenal androgen levels (17-hydroxyprogesterone, androstenedione) compared to baseline and placebo.

Secondary Objectives:
1. To assess the impact of Crinetide on quality of life measures using validated CAH-specific questionnaires.
2. To evaluate the pharmacokinetic profile of Crinetide and its relationship to clinical response.
3. To assess changes in body composition and metabolic parameters during Crinetide treatment.
4. To evaluate the effect of Crinetide on bone mineral density and bone turnover markers.
5. To assess patient-reported outcomes including fatigue, mood, and overall well-being.`,

        methodology: `Study Design: This Phase 2, randomized, double-blind, placebo-controlled, dose-ranging study will evaluate the safety and efficacy of Crinetide in adult patients with Congenital Adrenal Hyperplasia. The study employs a parallel-group design with three active dose arms and one placebo arm, with adaptive dose selection based on interim safety and efficacy data.

Patient Population: Eligible patients will be adults aged 18-65 years with genetically confirmed CAH due to 21-hydroxylase deficiency, currently receiving stable glucocorticoid replacement therapy for at least 6 months. Patients must have documented elevated 17-hydroxyprogesterone levels (>2x upper limit of normal) and demonstrate suboptimal disease control despite optimized conventional therapy. Key exclusion criteria include pregnancy, significant comorbidities, and use of investigational drugs within 30 days.

Treatment Regimen: Patients will be randomized 1:1:1:1 to receive subcutaneous Crinetide at doses of 0.5 mg, 1.0 mg, or 2.0 mg daily, or matching placebo, for 12 weeks. All patients will continue their background glucocorticoid therapy with dose adjustments permitted based on clinical response. The study includes a 4-week screening period, 12-week treatment period, and 4-week follow-up period.

Statistical Considerations: The study is powered to detect a 30% difference in morning cortisol normalization between active treatment and placebo arms, with 80% power and α=0.05. A total of 80 patients (20 per arm) will be enrolled to account for potential dropouts. Primary efficacy analysis will use ANCOVA with baseline cortisol as covariate, and safety analysis will include all randomized patients who receive at least one dose of study drug.`
    };

    // Simulate API delay
    await new Promise(resolve => setTimeout(resolve, 1500));
    
    return mockResponses[section];
}

async function insertTextIntoDocument(text, contentControlTag) {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                // Try to find the content control by tag
                const contentControls = context.document.contentControls.getByTag(contentControlTag);
                contentControls.load("items");
                
                await context.sync();
                
                if (contentControls.items.length > 0) {
                    // Insert text into the first content control with this tag
                    contentControls.items[0].insertText(text, "Replace");
                } else {
                    // If no content control found, insert at the beginning of the document
                    context.document.body.insertParagraph(text, "Start");
                }
                
                await context.sync();
                resolve();
            } catch (error) {
                reject(error);
            }
        });
    });
}

function showStatus(message, type = "loading") {
    const statusSection = document.getElementById("status");
    const statusText = document.getElementById("status-text");
    
    statusText.textContent = message;
    statusSection.style.display = "flex";
    
    // Remove existing classes
    statusSection.classList.remove("success", "error");
    
    // Add appropriate class
    if (type === "success") {
        statusSection.classList.add("success");
    } else if (type === "error") {
        statusSection.classList.add("error");
    }
}

function hideStatus() {
    const statusSection = document.getElementById("status");
    statusSection.style.display = "none";
}

function disableButtons(disabled) {
    const buttons = document.querySelectorAll(".ms-Button");
    buttons.forEach(button => {
        button.disabled = disabled;
    });
} 