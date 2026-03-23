let currentTemplate = null;
let templates = {};

Office.onReady(() => {
  loadTemplates();
  setupEventListeners();
  updateTemplateSelect();
  updateRulesDisplay();
});

function setupEventListeners() {
  document.getElementById('template-select').onchange = onTemplateSelect;
  document.getElementById('load-template-button').onclick = loadSelectedTemplate;
  document.getElementById('new-template-button').onclick = createNewTemplate;
  document.getElementById('save-template-button').onclick = saveCurrentTemplate;
  document.getElementById('delete-template-button').onclick = deleteCurrentTemplate;
  document.getElementById('format-button').onclick = formatDocument;
  document.getElementById('format-section-button').onclick = formatSelection;
}

function loadTemplates() {
  const stored = localStorage.getItem('resumeTemplates');
  if (stored) {
    templates = JSON.parse(stored);
  } else {
    // Create default template
    templates['Default'] = {
      name: 'Default',
      fontName: 'Calibri',
      fontSize: 11,
      lineSpacing: 14,
      headingKeywords: 'summary,professional summary,experience,education,skills,certifications,projects',
      bulletPatterns: '^(\\*|\\-|•|\\d+\\.)\\s+',
      datePattern: '\\b(\\d{4}|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\\b',
      dateAlignment: 'Right'
    };
    saveTemplates();
  }
}

function saveTemplates() {
  localStorage.setItem('resumeTemplates', JSON.stringify(templates));
}

function updateTemplateSelect() {
  const select = document.getElementById('template-select');
  select.innerHTML = '<option value="">Select a template...</option>';
  
  Object.keys(templates).forEach(name => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    select.appendChild(option);
  });
}

function onTemplateSelect() {
  const select = document.getElementById('template-select');
  const selectedName = select.value;
  
  if (selectedName && templates[selectedName]) {
    loadTemplateIntoEditor(templates[selectedName]);
    document.getElementById('template-editor').style.display = 'block';
    document.getElementById('save-template-button').style.display = 'inline-block';
    document.getElementById('delete-template-button').style.display = 'inline-block';
  } else {
    document.getElementById('template-editor').style.display = 'none';
    document.getElementById('save-template-button').style.display = 'none';
    document.getElementById('delete-template-button').style.display = 'none';
  }
}

function loadSelectedTemplate() {
  const select = document.getElementById('template-select');
  const selectedName = select.value;
  
  if (selectedName && templates[selectedName]) {
    currentTemplate = templates[selectedName];
    updateRulesDisplay();
    setStatus(`Template "${selectedName}" loaded.`);
  } else {
    setStatus('Please select a template first.');
  }
}

function createNewTemplate() {
  currentTemplate = {
    name: '',
    fontName: 'Calibri',
    fontSize: 11,
    lineSpacing: 14,
    headingKeywords: 'summary,professional summary,experience,education,skills,certifications,projects',
    bulletPatterns: '^(\\*|\\-|•|\\d+\\.)\\s+',
    datePattern: '\\b(\\d{4}|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\\b',
    dateAlignment: 'Right'
  };
  
  loadTemplateIntoEditor(currentTemplate);
  document.getElementById('template-editor').style.display = 'block';
  document.getElementById('save-template-button').style.display = 'inline-block';
  document.getElementById('delete-template-button').style.display = 'none'; // Can't delete unsaved template
  updateRulesDisplay();
  setStatus('New template created. Configure settings and save.');
}

function loadTemplateIntoEditor(template) {
  document.getElementById('template-name').value = template.name || '';
  document.getElementById('font-name').value = template.fontName || 'Calibri';
  document.getElementById('font-size').value = template.fontSize || 11;
  document.getElementById('line-spacing').value = template.lineSpacing || 14;
  document.getElementById('heading-keywords').value = template.headingKeywords || '';
  document.getElementById('bullet-patterns').value = template.bulletPatterns || '';
  document.getElementById('date-pattern').value = template.datePattern || '';
  document.getElementById('date-alignment').value = template.dateAlignment || 'Right';
}

function saveCurrentTemplate() {
  const name = document.getElementById('template-name').value.trim();
  
  if (!name) {
    setStatus('Please enter a template name.');
    return;
  }
  
  const template = {
    name: name,
    fontName: document.getElementById('font-name').value,
    fontSize: parseInt(document.getElementById('font-size').value),
    lineSpacing: parseInt(document.getElementById('line-spacing').value),
    headingKeywords: document.getElementById('heading-keywords').value,
    bulletPatterns: document.getElementById('bullet-patterns').value,
    datePattern: document.getElementById('date-pattern').value,
    dateAlignment: document.getElementById('date-alignment').value
  };
  
  templates[name] = template;
  currentTemplate = template;
  saveTemplates();
  updateTemplateSelect();
  updateRulesDisplay();
  setStatus(`Template "${name}" saved successfully.`);
}

function deleteCurrentTemplate() {
  const name = document.getElementById('template-name').value.trim();
  
  if (!name || !templates[name]) {
    setStatus('No template to delete.');
    return;
  }
  
  if (confirm(`Are you sure you want to delete the template "${name}"?`)) {
    delete templates[name];
    saveTemplates();
    updateTemplateSelect();
    document.getElementById('template-editor').style.display = 'none';
    document.getElementById('save-template-button').style.display = 'none';
    document.getElementById('delete-template-button').style.display = 'none';
    currentTemplate = null;
    updateRulesDisplay();
    setStatus(`Template "${name}" deleted.`);
  }
}

function updateRulesDisplay() {
  const display = document.getElementById('rules-display');
  
  if (!currentTemplate) {
    display.innerHTML = '<p>No template loaded. Select or create a template to see formatting rules.</p>';
    return;
  }
  
  const rules = `
    <h3>${currentTemplate.name} Template Rules</h3>
    <ul>
      <li><strong>Font:</strong> ${currentTemplate.fontName} ${currentTemplate.fontSize}pt</li>
      <li><strong>Line Spacing:</strong> ${currentTemplate.lineSpacing} points</li>
      <li><strong>Heading Keywords:</strong> ${currentTemplate.headingKeywords}</li>
      <li><strong>Bullet Patterns:</strong> ${currentTemplate.bulletPatterns}</li>
      <li><strong>Date Pattern:</strong> ${currentTemplate.datePattern}</li>
      <li><strong>Date Alignment:</strong> ${currentTemplate.dateAlignment}</li>
    </ul>
  `;
  
  display.innerHTML = rules;
}

function setStatus(message) {
  const status = document.getElementById('status');
  status.textContent = message;
}

async function formatDocument() {
  if (!currentTemplate) {
    setStatus('Please load a template first.');
    return;
  }
  
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clearFormatting();
      body.font.name = currentTemplate.fontName;
      body.font.size = currentTemplate.fontSize;
      body.paragraphFormat.spaceAfter = 0;
      body.paragraphFormat.lineSpacing = currentTemplate.lineSpacing;

      const paragraphs = body.paragraphs;
      paragraphs.load('items');
      await context.sync();

      // Parse heading keywords
      const headingRegex = new RegExp(`^(${currentTemplate.headingKeywords.split(',').map(k => k.trim()).join('|')})$`, 'i');
      
      // Parse bullet patterns
      const bulletRegex = new RegExp(currentTemplate.bulletPatterns);
      
      // Parse date pattern
      const dateRegex = new RegExp(currentTemplate.datePattern, 'i');

      for (let para of paragraphs.items) {
        const text = para.text.trim();

        // Apply heading styles
        if (headingRegex.test(text)) {
          para.style = 'Heading 2';
          para.font.bold = true;
        } else {
          para.style = 'Normal';
          para.font.bold = false;
        }

        // Apply bullet formatting
        if (bulletRegex.test(text)) {
          para.listItemType = 'Bullet';
        }

        // Apply date alignment
        if (dateRegex.test(text)) {
          para.alignment = currentTemplate.dateAlignment;
        } else {
          para.alignment = 'Left';
        }
      }

      // Remove extra blank paragraphs
      const paragraphs2 = body.paragraphs;
      paragraphs2.load('items');
      await context.sync();

      for (let i = paragraphs2.items.length - 1; i >= 0; i--) {
        const p = paragraphs2.items[i];
        if (p.text.trim() === '') {
          p.delete(true);
        }
      }

      await context.sync();
    });

    setStatus('Document formatted successfully using template "' + currentTemplate.name + '".');
  } catch (error) {
    console.error(error);
    setStatus('Error formatting document: ' + (error.message || error));
  }
}

async function formatSelection() {
  if (!currentTemplate) {
    setStatus('Please load a template first.');
    return;
  }
  
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.name = currentTemplate.fontName;
      selection.font.size = currentTemplate.fontSize;
      selection.paragraphFormat.lineSpacing = currentTemplate.lineSpacing;
      await context.sync();
    });
    setStatus('Selected text formatted successfully using template "' + currentTemplate.name + '".');
  } catch (error) {
    setStatus('Error formatting selection: ' + (error.message || error));
  }
}
