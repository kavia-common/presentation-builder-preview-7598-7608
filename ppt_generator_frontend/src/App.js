import React, { useCallback } from 'react';
import './App.css';
import './styles/theme.css';
import WizardLayout from './components/wizard/WizardLayout';
import { WizardProvider } from './state/wizardContext';
import { loadTemplateBundle, loadWizardSchema } from './services/schemaLoader';

// PUBLIC_INTERFACE
function App() {
  /** Entrypoint: loads schemas, provides wizard context, renders the wizard layout. */
  const loadSchemas = useCallback(async () => {
    const [wizardSchema, templateBundle] = await Promise.all([loadWizardSchema(), loadTemplateBundle()]);
    return { wizardSchema, ...templateBundle };
  }, []);

  return (
    <WizardProvider loadSchemas={loadSchemas}>
      <WizardLayout />
    </WizardProvider>
  );
}

export default App;
