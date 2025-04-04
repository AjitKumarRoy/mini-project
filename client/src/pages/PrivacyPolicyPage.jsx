import React from 'react';
import { Link } from 'react-router-dom'; // If you need navigation

const PrivacyPolicyPage = () => {
  return (
    <div className="bg-gray-100 py-12 sm:py-16 lg:py-20">
      <div className="container mx-auto px-4 sm:px-6 lg:px-8 max-w-3xl bg-white rounded-xl shadow-xl overflow-hidden">
        <div className="p-6 sm:p-8 lg:p-10">
          <header className="mb-8 text-center">
            <h1 className="text-3xl font-extrabold text-indigo-700 mb-2">Privacy Policy</h1>
            <p className="text-sm text-gray-500">Last Updated: 4/4/2025</p>
          </header>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Your Privacy Matters</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              Thank you for using Easy Sheets! We are committed to protecting your privacy. This policy explains how we collect, use, and share your information. We understand the importance of your personal data and strive to handle it with utmost care and transparency.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Information We Collect</h2>
            <p className="text-gray-700 leading-relaxed mb-4">We collect different types of information to provide and improve our services:</p>
            <ul className="list-disc pl-6 text-gray-700 leading-relaxed">
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Account Information:</strong> When you create an account, we may collect your name, email address, username, and password (stored securely). You may also provide additional optional information.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Usage Data:</strong> We collect information about how you interact with our app, including the features you use, the spreadsheets you create or edit, the time and duration of your sessions, and any errors you might encounter. This helps us understand user behavior and improve the app's functionality.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Device Information:</strong> We may collect details about the device you use to access our app, such as the device model, operating system version, unique device identifiers, IP address, and mobile network information. This information helps us optimize the app for different devices and troubleshoot issues.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Location Data (if applicable):</strong> If our app offers location-based features, we will explicitly ask for your permission before accessing your device's location. You can manage location permissions through your device settings at any time.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Content You Provide:</strong> This includes the data you enter into your spreadsheets, any files you upload, and any comments or collaborations you engage in within the app.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Communication Data:</strong> If you contact our support team, we may collect your name, email address, and the content of your communication to provide assistance.
              </li>
            </ul>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">How We Use Your Information</h2>
            <p className="text-gray-700 leading-relaxed mb-4">We use the collected information for various purposes to enhance your experience and improve our services:</p>
            <ul className="list-disc pl-6 text-gray-700 leading-relaxed">
              <li className="mb-3">To provide, operate, and maintain our app and its features.</li>
              <li className="mb-3">To personalize your experience and tailor content to your preferences.</li>
              <li className="mb-3">To communicate with you regarding your account, updates, security alerts, and support inquiries.</li>
              <li className="mb-3">To analyze app usage patterns and trends to understand how users interact with our platform and identify areas for improvement.</li>
              <li className="mb-3">To develop new features, functionalities, and services.</li>
              <li className="mb-3">To troubleshoot technical issues and provide customer support.</li>
              <li className="mb-3">To monitor and prevent fraudulent activities, misuse of our app, and other security incidents.</li>
              <li className="mb-3">To comply with legal obligations and enforce our terms of service.</li>
              <li className="mb-3">With your consent, for other purposes as disclosed to you.</li>
            </ul>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">How We Share Your Information</h2>
            <p className="text-gray-700 leading-relaxed mb-4">We are committed to protecting your personal information. We may share your information with:</p>
            <ul className="list-disc pl-6 text-gray-700 leading-relaxed">
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Service Providers:</strong> We engage trusted third-party service providers to assist us with various functions, such as hosting, data analysis, email delivery, customer support, and payment processing. These providers have access to your information only to perform these tasks on our behalf and are obligated not to disclose or use it for any other purpose.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Legal Compliance:</strong> We may disclose your information if required to do so by law or in response to valid legal processes, such as subpoenas, court orders, or government requests.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Business Transfers:</strong> In the event of a merger, acquisition, reorganization, sale of assets, or bankruptcy, your information may be transferred to the acquiring entity or third party as part of the transaction. We will notify you via email and/or a prominent notice on our app of any change in ownership or uses of your personal information.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">With Your Consent:</strong> We may share your information with third parties when we have your explicit consent to do so.
              </li>
              <li className="mb-3">
                <strong className="font-semibold text-gray-800">Aggregated and Anonymized Data:</strong> We may share aggregated and anonymized data that does not personally identify you with third parties for various purposes, including analytics, research, and marketing.
              </li>
            </ul>
            <p className="text-gray-700 leading-relaxed mt-4">
              <strong className="font-semibold text-gray-800">We do not sell your personal information to third parties for their direct marketing purposes.</strong>
            </p>
          </section>

          {/* Add more sections from the previous Privacy Policy content here, formatted for better readability */}

          <footer className="mt-12 text-center text-gray-500">
            <p className="mb-2">© {new Date().getFullYear()} Easy Sheets. All rights reserved.</p>
            <Link to="mailto:privacy@yourapp.com" className="text-indigo-500 hover:underline">Contact Us Regarding Privacy</Link>
          </footer>
        </div>
      </div>
    </div>
  );
};

export default PrivacyPolicyPage;