# WPS Office DeepSeek AI Assistant

![WPS-DeepSeek Integration](https://via.placeholder.com/800x400.png?text=WPS+Office+%26+DeepSeek+AI+Integration)

A smart office add-in that integrates DeepSeek's AI capabilities directly into WPS Office Professional. Developed by Prof. Lingzhang meng, from Institut for Cardiovascular Diseases, Guangxi Academy of Medical Sciences & The People's Hospital of Guangxi Zhuang Autonomous Region (广西医学科学院&广西壮族自治区人民医院 心血管疾病研究所 孟令章 教授）.

## Key Features ✨

- **Seamless AI Integration**
  - Direct access to DeepSeek's advanced NLP services
  - Real-time text analysis within WPS documents
  - API-based communication with enterprise-level security

- **Smart Text Processing**
  - Automatic markdown cleansing
  - Intelligent text formatting with color-coding
  - Structured output for suggestions/cautions
  - Multi-level content organization

- **Professional Enhancement**
  - Medical research-optimized processing
  - Academic writing support
  - Technical document refinement
  - Multi-language compatibility

## Key Advantages 🚀

1. **Deep Context Understanding**
   - Leverages DeepSeek's 160B parameter model
   - Maintains document context awareness
   - Supports complex medical terminology

2. **Workflow Integration**
   - Direct in-document operation
   - One-click text analysis (Alt+F8)
   - Format-preserving processing

3. **Enhanced Security**
   - API key authentication
   - Local text preprocessing
   - Secure HTTPS communication

4. **Smart Formatting Engine**
   - Automatic list recognition
   - Color-coded suggestions (Green) & cautions (Red)
   - Multi-level content segmentation
   - Clean markdown-to-plaintext conversion

## Installation 📥

1. Clone repository or download `DeepSeek_WPS_Addin.bas`
2. Open WPS Office Developer Tools (Alt+F11)
3. Import module to your document/project
4. Insert your DeepSeek API key in line 47
5. Enable Macros:
   Open WPS pro Writer: Go to File > Options > Trust Center > Macro Security
   Select "Enable all macros" (temporarily for setup)
6. Insert the Macro:
   Press Alt + F11 to open VBA editor, Right-click your document in the Project Explorer, Select Insert > Module, Paste the provided code
8. Set API Key:
   In the code, locate api_key = "Your_API_Key_strings", Replace "Your_API_Key_strings" with your actual DeepSeek API key
9. Assign Shortcut:
   Go to Tools > Customize Keyboard, Find and assign a shortcut to the 智慧办公 macro
10. Use the Macro:
   Select text in your document, Run the macro (either through the macro dialog or assigned shortcut)

The AI response will appear below your selection with formatted suggestions

```vba
apply your API key following DeepSeek guidlines from its official website.
api_key = "Your_API_Key_String" ' Replace with actual key

<img width="632" alt="Weixin Image_20250304110142" src="https://github.com/user-attachments/assets/c919cd47-2cf6-4567-a2ca-08f289d990fa" />

