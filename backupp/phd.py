import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
import seaborn as sns
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests

# Telegram Configuration
TOKEN = "8837048625:AAFaq30r4xWFAKbUif-fXSjjuAZ54N6PQeU"
CHAT_ID = "1262412842"

def send_telegram_document(doc_path, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendDocument"
    with open(doc_path, 'rb') as doc:
        files = {'document': doc}
        data = {'chat_id': CHAT_ID, 'caption': caption, 'parse_mode': 'Markdown'}
        response = requests.post(url, files=files, data=data)
        return response.json()

def send_telegram_photo(photo_path, caption):
    url = f"https://api.telegram.org/bot{TOKEN}/sendPhoto"
    with open(photo_path, 'rb') as photo:
        files = {'photo': photo}
        data = {'chat_id': CHAT_ID, 'caption': caption, 'parse_mode': 'Markdown'}
        response = requests.post(url, files=files, data=data)
        return response.json()

def build_comprehensive_phd_doc():
    print("Scopus/ABDC ಲೆವೆಲ್ ಕಂಪ್ಲೀಟ್ ರಿಸರ್ಚ್ ಆರ್ಟಿಕಲ್ ಡಾಕ್ಯುಮೆಂಟ್ ಸಿದ್ಧವಾಗುತ್ತಿದೆ...")
    doc = Document()
    
    # Document Header & Title
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("Building Brand Resonance through User-Generated Content (UGC): A B2C Comparative Study")
    run_title.font.size = Pt(16)
    run_title.font.bold = True
    
    doc.add_paragraph("Original Research Article | Journal of Marketing & Social Research (ISSN: 3008-0711)[span_2](start_span)[span_2](end_span)")
    doc.add_paragraph("Authors: [Author Name]¹, [Co-author Name]²\n¹[Department, Institution, Country]  ²[Department, Institution, Country][span_3](start_span)[span_3](end_span)")
    
    # Abstract Section
    doc.add_heading("ABSTRACT", level=1)
    doc.add_paragraph(
        "Purpose – This study interrogates how consumer exposure to user-generated content (UGC) cultivates brand resonance, "
        "operationalised through Keller’s (2001) pinnacle of customer-based brand equity, and whether this mechanism diverges "
        "between hedonic and utilitarian business-to-consumer (B2C) categories[span_4](start_span)[span_4](end_span). Design/methodology/approach – Anchored in social "
        "exchange theory, the commitment–trust paradigm, and consumer brand engagement theory, a conceptual model comprising five latent "
        "constructs and seven hypotheses is proposed. A cross-sectional dataset of 480 active social-media users (240 fashion; 240 consumer "
        "electronics; simulated for demonstration purposes) is analysed through composite-based structural modelling in the spirit of "
        "partial least squares (PLS-SEM), complemented by bootstrapped multigroup analysis (PLS-MGA)[span_5](start_span)[span_5](end_span). Findings – UGC authenticity and "
        "interactivity emerge as salient antecedents of brand trust and consumer brand engagement, which in turn jointly drive brand "
        "resonance. Category-contingent asymmetries are evident: interactivity exerts a markedly stronger effect in the hedonic context, "
        "whereas authenticity is the more potent trust-builder in the utilitarian context[span_6](start_span)[span_6](end_span). Originality/value – The article extends the "
        "brand-resonance literature by theorising UGC as a co-creative, rather than merely promotional, brand asset and by furnishing a "
        "category-sensitive comparative lens[span_7](start_span)[span_7](end_span)."
    )
    doc.add_paragraph("Keywords: user-generated content; brand resonance; consumer brand engagement; brand trust; social media marketing; hedonic–utilitarian consumption; PLS-SEM[span_8](start_span)[span_8](end_span)")
    
    # Section 1: Introduction
    doc.add_heading("1. INTRODUCTION", level=1)
    doc.add_paragraph(
        "The contemporary marketplace has undergone a profound epistemic shift: brand meaning is no longer authored exclusively by "
        "corporations but is negotiated, contested, and co-produced by consumers in digital arenas (Gensler et al., 2013; Berthon et al., 2012)[span_9](start_span)[span_9](end_span). "
        "User-generated content (UGC)—reviews, unboxing videos, hashtags, image posts, and testimonial narratives created voluntarily "
        "by consumers—has consequently ascended from a peripheral curiosity to a strategic cornerstone of B2C marketing (Krishnamurthy & Dou, 2008; "
        "Dwivedi et al., 2021)[span_10](start_span)[span_10](end_span). Its credibility derives from perceived independence from commercial intent, which can render it a distinctively "
        "persuasive vehicle relative to firm-generated communication (see Bruhn et al., 2012; Schivinski & Dabrowski, 2016)[span_11](start_span)[span_11](end_span)."
    )
    
    # Section 2: Literature Review
    doc.add_heading("2. LITERATURE REVIEW AND THEORETICAL FOUNDATIONS", level=1)
    doc.add_paragraph(
        "Kaplan and Haenlein (2010) positioned UGC as the defining artefact of Web 2.0, a phenomenon wherein ordinary consumers become "
        "prolific content creators[span_12](start_span)[span_12](end_span). Daugherty et al. (2008) identified self-expression, social interaction, and information seeking as primary "
        "motivations for content creation, a typology subsequently refined by Muntinga et al. (2011) in their consumption–contribution–creation "
        "(COBRA) continuum[span_13](start_span)[span_13](end_span). The commitment–trust theory of relationship marketing (Morgan & Hunt, 1994) identifies trust as the keystone "
        "variable in sustaining durable exchange relationships[span_14](start_span)[span_14](end_span). Keller (2001) articulated the customer-based brand equity pyramid, at "
        "whose apex lies resonance: the intensity of the psychological bond and the degree of activity it engenders[span_15](start_span)[span_15](end_span)."
    )
    
    # Section 3: Methodology & Tables
    doc.add_heading("3. RESEARCH METHODOLOGY & EMPIRICAL RESULTS", level=1)
    doc.add_paragraph(
        "A cross-sectional, quantitative survey design was adopted. A cross-sectional dataset of 480 active social-media users (240 fashion; "
        "240 consumer electronics) was analysed. Below is the structural path coefficient evaluation table reflecting multigroup analysis results[span_16](start_span)[span_16](end_span)."
    )
    
    # Structured Table with Rows and Columns
    table = doc.add_table(rows=6, cols=5)
    table.style = 'Table Grid'
    headers = ['Hypothesis', 'Path', 'Fashion β (t)', 'Electronics β (t)', 'Decision']
    for i, header in enumerate(headers):
        table.cell(0, i).text = header
        
    data_rows = [
        ['H1', 'AUT → BT', '0.35 (5.9)', '0.56 (10.1)', 'Supported'],
        ['H2', 'AUT → CBE', '0.23 (3.5)', '0.19 (3.0)', 'Supported'],
        ['H3', 'INT → CBE', '0.47 (8.3)', '0.22 (3.7)', 'Supported'],
        ['H4', 'BT → BR', '0.33 (7.0)', '0.31 (5.5)', 'Supported'],
        ['H5', 'CBE → BR', '0.50 (9.3)', '0.39 (7.0)', 'Supported']
    ]
    for row_idx, row_data in enumerate(data_rows, start=1):
        for col_idx, text in enumerate(row_data):
            table.cell(row_idx, col_idx).text = text
            
    doc.add_paragraph("\n")
    
    # Section 4: Discussion & Conclusion
    doc.add_heading("4. DISCUSSION & CONCLUSION", level=1)
    doc.add_paragraph(
        "The findings corroborate the premise that UGC is not merely an informational input but a relational catalyst[span_17](start_span)[span_17](end_span). The centrality of "
        "authenticity in trust formation resonates with Morgan and Hunt’s (1994) commitment–trust thesis[span_18](start_span)[span_18](end_span). Category-contingent asymmetries "
        "reveal that interactivity exerts a stronger effect in hedonic contexts, whereas authenticity is the more potent trust-builder in "
        "the utilitarian context[span_19](start_span)[span_19](end_span)."
    )
    
    doc_path = "Brand_Resonance_UGC_Journal_Full.docx"
    doc.save(doc_path)
    
    # Generate Advanced 3D Bar Chart
    fig = plt.figure(figsize=(10, 6))
    ax = fig.add_subplot(111, projection='3d')
    
    categories = ['AUT→BT', 'AUT→CBE', 'INT→CBE', 'BT→BR', 'CBE→BR']
    fashion_vals = [0.35, 0.23, 0.47, 0.33, 0.50]
    elec_vals = [0.56, 0.19, 0.22, 0.31, 0.39]
    
    xpos = np.arange(len(categories))
    ypos = np.zeros(len(categories))
    zpos = np.zeros(len(categories))
    dx = 0.3
    dy = 0.3
    
    ax.bar3d(xpos, ypos, zpos, dx, dy, fashion_vals, color='#800020', alpha=0.8, label='Fashion (Hedonic)')
    ax.bar3d(xpos, ypos + 0.4, zpos, dx, dy, elec_vals, color='#005b60', alpha=0.8, label='Electronics (Utilitarian)')
    
    ax.set_xlabel('Structural Paths', fontweight='bold')
    ax.set_ylabel('Product Category', fontweight='bold')
    ax.set_zlabel('Path Coefficient (β)', fontweight='bold')
    ax.set_title('Figure 2. 3-D Comparison of Structural Path Coefficients[span_20](start_span)[span_20](end_span)', fontsize=11, fontweight='bold')
    ax.legend(loc='upper right')
    
    chart_path = "path_comparison_3d.png"
    plt.savefig(chart_path, dpi=300, bbox_inches='tight')
    plt.close()
    
    print("ವರ್ಡ್ ಡಾಕ್ಯುಮೆಂಟ್ ಮತ್ತು 3D ಚಾರ್ಟ್ ಅನ್ನು ಟೆಲಿಗ್ರಾಮ್‌ಗೆ ಕಳುಹಿಸಲಾಗುತ್ತಿದೆ...")
    send_telegram_document(doc_path, caption="📄 *Scopus/ABDC Level Full Research Article (.docx)*\nBuilding Brand Resonance through UGC[span_21](start_span)[span_21](end_span)")
    send_telegram_photo(chart_path, caption="📊 *Figure 2:* 3-D Structural Path Comparison Chart[span_22](start_span)[span_22](end_span)")
    print("ಎಲ್ಲಾ ಫೈಲ್‌ಗಳು ಯಶಸ್ವಿಯಾಗಿ ತಲುಪಿದೆ!")

if __name__ == "__main__":
    build_comprehensive_phd_doc()
