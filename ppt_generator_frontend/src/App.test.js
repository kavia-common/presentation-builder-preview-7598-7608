import { render, screen } from '@testing-library/react';
import App from './App';

test('renders presentation builder shell', async () => {
  render(<App />);

  // Initial state shows the loading skeleton while schemas are fetched.
  expect(screen.getByText(/loading schemas/i)).toBeInTheDocument();

  // Once loaded, we should see the app header.
  expect(await screen.findByText(/presentation builder/i)).toBeInTheDocument();
});
