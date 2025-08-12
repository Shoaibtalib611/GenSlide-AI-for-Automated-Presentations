/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    // Hide sideload message and show app body if present
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) sideloadMsg.style.display = "none";
    const appBody = document.getElementById("app-body");
    if (appBody) appBody.style.display = "flex";

    // Attach event listeners for AI slide generation
    document.getElementById("generateSlides").onclick = generateSlides;
    document.getElementById("insertSlides").onclick = insertSlides;
  }
});

let generatedSlides = [];

async function generateSlides() {
  const topic = document.getElementById("topic").value;
  const slideCount = parseInt(document.getElementById("slideCount").value, 10);
  const layout = document.getElementById("layout").value;
  const notification = document.getElementById("notification");

  if (!topic || !slideCount || slideCount < 1) {
    notification.innerText = "Please enter a topic and a valid number of slides.";
    notification.style.display = "block";
    return;
  }
  notification.innerText = "Generating slides using AI...";
  notification.style.display = "block";

  // Prepare prompt for Gemini API
  const prompt = `Generate content for a PowerPoint presentation on the topic: '${topic}'. Create ${slideCount} slides. For each slide, provide a title and bullet points. Layout: ${layout}. Return as a JSON array: [{title: string, bullets: string[]}]`;

  try {
    const apiKey = "AIzaSyB-pQRhd4s4wPyiCpBuH3zTpDAVjZOcEJM";
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }]
        })
      }
    );
    const data = await response.json();
    // Parse Gemini response
    let slidesJson = null;
    let aiText = null;
    if (data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts && data.candidates[0].content.parts[0].text) {
      aiText = data.candidates[0].content.parts[0].text;
      // Try to extract JSON from the AI response, even if it is surrounded by text/code block
      const jsonMatch = aiText.match(/\[.*\]/s);
      if (jsonMatch) {
        try {
          slidesJson = JSON.parse(jsonMatch[0]);
        } catch (e) {
          notification.innerText = "AI response JSON could not be parsed. Try again.";
          return;
        }
      } else {
        notification.innerText = "AI did not return JSON slide data. Try again.";
        return;
      }
    }
    if (!slidesJson || !Array.isArray(slidesJson)) {
      notification.innerText = "AI did not return valid slide data.";
      return;
    }
    generatedSlides = slidesJson;
    notification.innerText = `Generated ${generatedSlides.length} slides. Click 'Insert Slides' to add them to PowerPoint.`;
    document.getElementById("insertSlides").style.display = "block";
  } catch (err) {
    notification.innerText = "Error generating slides: " + err.message;
  }
}

async function insertSlides() {
  const notification = document.getElementById("notification");
  if (!generatedSlides.length) {
    notification.innerText = "No slides to insert. Please generate slides first.";
    notification.style.display = "block";
    return;
  }
  try {
    await Office.onReady();
    await PowerPoint.run(async function(context) {
      let slides = context.presentation.slides;
      // Prepare slide content as a list of lists (each slide: [title, ...bullets])
      let slideContents = generatedSlides.map(slide => {
        let content = [];
        if (slide.title) content.push(slide.title);
        if (slide.bullets && Array.isArray(slide.bullets)) {
          content = content.concat(slide.bullets);
        }
        return content;
      });
      // Insert content into slides
      for (let i = 0; i < slideContents.length; i++) {
        let contentList = slideContents[i];
        let slide;
        if (i === 0 && slides.count > 0) {
          slide = slides.getItemAt(0);
        } else {
          slide = slides.add();
          await context.sync();
          // Move reference to the newly added slide
          slide = slides.getItemAt(i);
          await context.sync();
        }
        let shapes = slide.shapes;
        // Insert title as a textbox
        if (contentList.length > 0) {
          let titleShape = shapes.addTextBox(contentList[0]);
          titleShape.left = 50;
          titleShape.top = 50;
          titleShape.width = 600;
          titleShape.height = 50;
        }
        // Insert bullets as a single textbox
        if (contentList.length > 1) {
          let bullets = contentList.slice(1).map(b => `• ${b}`).join("\n");
          let bulletShape = shapes.addTextBox(bullets);
          bulletShape.left = 50;
          bulletShape.top = 120;
          bulletShape.width = 600;
          bulletShape.height = 300;
        }
        await context.sync();
        await new Promise(resolve => setTimeout(resolve, 400));
      }
    });
    notification.innerText = "All slides inserted!";
    document.getElementById("insertSlides").style.display = "none";
  } catch (err) {
    notification.innerText = "Error inserting slides: " + err.message;
  }
}

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}
