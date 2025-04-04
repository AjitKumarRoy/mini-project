import React from 'react'
import { Routes, Route } from 'react-router-dom'
import Login from './pages/Login'
import Dashboard from './pages/Dashboard'
import CreateSpreadsheet from './pages/CreateSpreadsheet'
import SheetDetail from './pages/SheetDetail'
import PrivateRoute from './components/PrivateRoute' // for protected routes
import PrivacyPolicyPage from './pages/PrivacyPolicyPage';
import TermsOfServicePage from './pages/TermsOfServicePage';
import Contact from './pages/Contact';
import About from './pages/About';
import Layout from './components/Layout'
import './App.css'

function App() {
  return (
    <Routes>
      {/* Public routes */}
      <Route path="/login" element={<Login />} />

      {/* Protected routes */}
      <Route
        path="/"
        element={
          <PrivateRoute>
            <Layout />
          </PrivateRoute>
        }
      >
        <Route index element={<Dashboard />} />
        <Route path="create-spreadsheet" element={<CreateSpreadsheet />} />
        <Route path="sheets/:sheetId" element={<SheetDetail />} />
        <Route path="privacy-policy" element={<PrivacyPolicyPage />} />
        <Route path="terms-of-service" element={<TermsOfServicePage />} />
        <Route path="contact" element={<Contact />} />
        <Route path="about" element={<About />} />
      </Route>
    </Routes>
  )
}

export default App
