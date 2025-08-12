# GenSlide: AI for Automated Presentations

GenSlide is an AI-powered Microsoft PowerPoint add-in designed to automatically generate professional presentations from user input. Using Google’s Gemini API, it creates well-structured slides with relevant text, images, and formatting in seconds. The tool enables users to interact through a chatbot interface, allowing customization and refinement of slides in real time. By automating slide creation, GenSlide saves time, enhances productivity, and ensures visually engaging presentations for academic, corporate, and creative use cases.

![GenSlide Preview](https://github.com/Shoaibtalib611/GenSlide-AI-for-Automated-Presentations/blob/main/Assets/preview.jpg)
![GenSlide UI](https://github.com/Shoaibtalib611/GenSlide-AI-for-Automated-Presentations/blob/main/Assets/Screenshot%202025-08-12%20193244.jpg)

---

## 🚀 Features
- AI-powered slide generation from plain text  
- Integration with Google Gemini API  
- Real-time chatbot customization  
- Automated design and layout suggestions  
- Easy deployment as a PowerPoint add-in  

---

## 📦 Installation

### Prerequisites
- Node.js (v16+ recommended)  
- npm  
- Yeoman & Office Add-in Generator  
- Microsoft Office 365 (with PowerPoint)  

```bash
# Install Yeoman and Office Add-in Generator
npm install -g yo generator-office

# Create a new Office Add-in project
yo office
Select:

Project type: Office Add-in Task Pane project

Script type: JavaScript / TypeScript (your choice)

Office application: PowerPoint

💻 Development
bash
Copy
Edit
# Install dependencies
npm install

# Start local server and sideload add-in into PowerPoint
npm start
🛑 Stop Development Server
bash
Copy
Edit
npm stop
📜 Usage
Open PowerPoint.

Load the GenSlide add-in (sideloaded automatically via npm start).

Enter your topic or prompt.

Review and customize AI-generated slides.

Save or export your presentation.

📌 Notes
Requires a valid Google Gemini API key.

Ensure Office is configured to allow sideloading add-ins.
