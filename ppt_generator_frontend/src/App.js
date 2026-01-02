import React, { useCallback } from 'react';
import './App.css';
import './styles/theme.css';
import WizardLayout from './components/wizard/WizardLayout';
import { WizardProvider } from './state/wizardContext';
import { loadTemplateModel, loadWizardSchema } from './services/schemaLoader';

// PUBLIC_INTERFACE
function App() {
  /** Entrypoint: loads schemas, provides wizard context, renders the wizard layout. */
  const loadSchemas = useCallback(async () => {
    const [wizardSchema, templateModel] = await Promise.all([loadWizardSchema(), loadTemplateModel()]);
    return { wizardSchema, templateModel };
  }, []);

  return (
    <WizardProvider loadSchemas={loadSchemas}>
      <WizardLayout />
    </WizardProvider>
  );
}

export default App;
