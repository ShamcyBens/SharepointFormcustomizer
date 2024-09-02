import * as React from 'react';
import { useState, useEffect } from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';

import styles from './DynamicForm.module.scss';

export interface IField {
  id: number;
  name: string;
  type: string;
  options?: string[];
}

export interface IDynamicFormProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

const LOG_SOURCE: string = 'DynamicForm';

const DynamicForm: React.FC<IDynamicFormProps> = ({ context, displayMode, onSave, onClose }) => {
  const [fields, setFields] = useState<IField[]>([]);
  const [templateLoaded, setTemplateLoaded] = useState<boolean>(false);

  useEffect(() => {
    const loadTemplate = async (): Promise<void> => {
      try {
        const itemId = context.itemId;  // Get the current item ID from the context
        const listUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Business')/items(${itemId})`;
        const response = await fetch(listUrl, {
          method: 'GET',
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        });
        const itemData = await response.json();
        const templateId = itemData.d.TemplateId; // Assume the Template ID field is named TemplateId

        const templateUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Business')/items(${templateId})`;
        const templateResponse = await fetch(templateUrl, {
          method: 'GET',
          headers: {
            'Accept': 'application/json;odata=verbose'
          }
        });
        const templateData = await templateResponse.json();
        setFields(templateData.d.fields);
        setTemplateLoaded(true);
      } catch (error) {
        Log.error(LOG_SOURCE, error);
      }
    };

    if (displayMode === FormDisplayMode.New || displayMode === FormDisplayMode.Edit) {
      loadTemplate().catch(error => Log.error(LOG_SOURCE, error));
    }
  }, [context, displayMode]);

  const addField = (type: string): void => {
    const newField: IField = { id: Date.now(), type, name: '', options: [] };
    setFields([...fields, newField]);
  };

  const updateField = (id: number, property: string, value: unknown): void => {
    const updatedFields = fields.map(field => field.id === id ? { ...field, [property]: value } : field);
    setFields(updatedFields);
  };

  const saveTemplate = async (): Promise<void> => {
    const templateName = prompt('Enter template name:');
    if (!templateName) return;

    const template = { Title: templateName, fields };
    const listUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Business')/items`;
    await fetch(listUrl, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify(template)
    }).catch(error => Log.error(LOG_SOURCE, error));
    alert('Template saved successfully');
  };

  const submitForm = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();
    const formData = new FormData(event.target as HTMLFormElement);
    const item: { [key: string]: unknown } = {};
    formData.forEach((value, key) => {
      item[key] = value;
    });

    const listUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Business')/items`;
    await fetch(listUrl, {
      method: 'POST',
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose'
      },
      body: JSON.stringify(item)
    }).catch(error => Log.error(LOG_SOURCE, error));
    alert('Form submitted successfully');
  };

  return (
    <div className={styles.dynamicForm}>
      {displayMode === FormDisplayMode.New || displayMode === FormDisplayMode.Edit ? (
        <form onSubmit={submitForm}>
          {templateLoaded ? (
            fields.map(field => (
              <div key={field.id}>
                {field.type === 'text' ? (
                  <input type="text" name={field.name} placeholder={field.name} />
                ) : (
                  <select name={field.name}>
                    {field.options?.map(option => (
                      <option key={option} value={option}>{option}</option>
                    ))}
                  </select>
                )}
              </div>
            ))
          ) : (
            <div>Loading template...</div>
          )}
          <button type="submit">Submit</button>
        </form>
      ) : (
        <div>
          <button onClick={() => addField('text')}>Add Text Field</button>
          <button onClick={() => addField('choice')}>Add Choice Field</button>
          <div id="fields">
            {fields.map(field => (
              <div key={field.id}>
                <input type="text" placeholder="Field Name" value={field.name}
                  onChange={(e) => updateField(field.id, 'name', e.target.value)} />
                {field.type === 'choice' && <textarea placeholder="Options (comma separated)"
                  onChange={(e) => updateField(field.id, 'options', e.target.value.split(','))} />}
              </div>
            ))}
          </div>
          <button onClick={saveTemplate}>Save Template</button>
        </div>
      )}
    </div>
  );
};

export default DynamicForm;
