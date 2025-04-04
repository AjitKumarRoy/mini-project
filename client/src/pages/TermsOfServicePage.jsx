import React from 'react';
import { Link } from 'react-router-dom'; // If you need navigation

const TermsOfServicePage = () => {
  return (
    <div className="bg-gray-100 py-12 sm:py-16 lg:py-20">
      <div className="container mx-auto px-4 sm:px-6 lg:px-8 max-w-3xl bg-white rounded-xl shadow-xl overflow-hidden">
        <div className="p-6 sm:p-8 lg:p-10">
          <header className="mb-8 text-center">
            <h1 className="text-3xl font-extrabold text-indigo-700 mb-2">Terms of Service</h1>
            <p className="text-sm text-gray-500">Last Updated: 4/4/2025</p>
          </header>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Welcome to Easy Sheets!</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              These Terms of Service ("Terms") govern your access to and use of the [mobile application/service] (referred to as the "App"). By accessing or using the App, you signify your agreement to be bound by these Terms. Please read them carefully.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Acceptance of Terms</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              By using the App, you acknowledge that you have read, understood, and agree to be bound by these Terms of Service, as well as our <Link to="/privacy-policy" className="text-indigo-500 hover:underline">Privacy Policy</Link>, which is incorporated herein by reference. If you do not agree to all of these terms, you are expressly prohibited from using the App and must discontinue use immediately.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Use of the App</h2>
            <p className="text-gray-700 leading-relaxed mb-4">
              You agree to use the App solely for lawful purposes and in a manner that respects the rights of others and does not interfere with their use and enjoyment of the App. You are responsible for all activity that occurs under your account and for maintaining the security of your account credentials.
            </p>
            <p className="text-gray-700 leading-relaxed mb-4">You must not:</p>
            <ul className="list-disc pl-6 text-gray-700 leading-relaxed">
              <li className="mb-3">Use the App in any way that could damage, disable, overburden, or impair the App or our servers or networks.</li>
              <li className="mb-3">Attempt to gain unauthorized access to any portion of the App, other user accounts, or computer systems or networks connected to the App through hacking, password mining, or any other means.</li>
              <li className="mb-3">Engage in any form of data mining, scraping, or extraction of content or information from the App without our express written consent.</li>
              <li className="mb-3">Upload or transmit any viruses, worms, malware, Trojan horses, or other harmful or disruptive code, files, or programs.</li>
              <li className="mb-3">Infringe upon or violate our intellectual property rights or the intellectual property rights of others.</li>
              <li className="mb-3">Use the App to transmit unsolicited or unauthorized advertising, promotional materials, spam, or any other form of solicitation.</li>
              <li className="mb-3">Impersonate any person or entity or falsely state or otherwise misrepresent your affiliation with a person or entity.</li>
              <li className="mb-3">Violate any applicable federal, state, local, or international law or regulation.</li>
            </ul>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">User Accounts</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              If you create an account on the App, you are responsible for maintaining the confidentiality of your username and password. You agree to provide accurate, current, and complete information during the registration process and to update such information to keep it accurate, current, and complete. You must notify us immediately of any unauthorized access to or use of your account or any other breach of security. We reserve the right to suspend or terminate your account at our sole discretion, without notice, for any reason, including but not limited to a violation of these Terms.
            </p>
          </section>

          {/* Add more sections from the previous Terms of Service content here, formatted for better readability */}

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Intellectual Property</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              The App and its original content, features, and functionality (including but not limited to text, graphics, logos, icons, images, audio clips, video clips, data compilations, software, and code) are owned by [Your Company Name] or its licensors and are protected by copyright, trademark, patent, trade secret, and other intellectual property or proprietary rights laws. You may not modify, reproduce, distribute, create derivative works of, publicly display, or in any way exploit any of the content or materials on the App in whole or in part except as expressly provided in these Terms or with our prior written consent.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Disclaimer of Warranties</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              THE APP IS PROVIDED ON AN "AS IS" AND "AS AVAILABLE" BASIS WITHOUT ANY WARRANTIES OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, NON-INFRINGEMENT, AND ANY WARRANTIES ARISING OUT OF COURSE OF DEALING OR USAGE OF TRADE. [Your Company Name] DOES NOT WARRANT THAT THE APP WILL BE UNINTERRUPTED, ERROR-FREE, SECURE, OR FREE OF VIRUSES OR OTHER HARMFUL COMPONENTS. YOU AGREE THAT YOUR USE OF THE APP IS AT YOUR SOLE RISK.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Limitation of Liability</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, IN NO EVENT SHALL [Your Company Name], ITS AFFILIATES, DIRECTORS, OFFICERS, EMPLOYEES, AGENTS, SUPPLIERS, OR LICENSORS BE LIABLE FOR ANY INDIRECT, INCIDENTAL, SPECIAL, CONSEQUENTIAL, OR PUNITIVE DAMAGES (INCLUDING WITHOUT LIMITATION DAMAGES FOR LOSS OF PROFITS, DATA, USE, GOODWILL, OR OTHER INTANGIBLE LOSSES) ARISING OUT OF OR RELATING TO YOUR ACCESS TO OR USE OF, OR INABILITY TO ACCESS OR USE, THE APP, WHETHER BASED ON WARRANTY, CONTRACT, TORT (INCLUDING NEGLIGENCE), STATUTE, OR ANY OTHER LEGAL THEORY, WHETHER OR NOT WE HAVE BEEN INFORMED OF THE POSSIBILITY OF SUCH DAMAGE, AND EVEN IF A REMEDY SET FORTH HEREIN IS FOUND TO HAVE FAILED OF ITS ESSENTIAL PURPOSE. OUR AGGREGATE LIABILITY TO YOU FOR ALL CLAIMS ARISING OUT OF OR RELATING TO THE APP SHALL NOT EXCEED THE AMOUNT YOU PAID, IF ANY, TO US FOR THE USE OF THE APP IN THE TWELVE (12) MONTHS PRIOR TO THE ACT GIVING RISE TO THE LIABILITY.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Governing Law</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              These Terms shall be governed by and construed in accordance with the laws of [Your Country/State], without regard to its conflict of law provisions. You agree to submit to the exclusive jurisdiction of the courts located in [Your Country/State] to resolve any dispute arising out of or relating to these Terms or the App.
            </p>
          </section>

          <section className="mb-10">
            <h2 className="text-2xl font-semibold text-gray-800 mb-4">Changes to These Terms</h2>
            <p className="text-gray-700 leading-relaxed mb-6">
              We reserve the right to modify or revise these Terms of Service at any time in our sole discretion. We will notify you of any material changes by posting the updated Terms on the App and updating the "Last Updated" date at the top. Your continued use of the App after the posting of revised Terms constitutes your acceptance of the changes. It is your responsibility to review these Terms periodically for any updates or modifications.
            </p>
          </section>

          <footer className="mt-12 text-center text-gray-500">
            <p className="mb-2">© {new Date().getFullYear()} Easy Sheets. All rights reserved.</p>
            <Link to="mailto:terms@yourapp.com" className="text-indigo-500 hover:underline">Contact Us Regarding Terms</Link>
          </footer>
        </div>
      </div>
    </div>
  );
};

export default TermsOfServicePage;