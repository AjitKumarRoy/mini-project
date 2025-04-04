import React, { useState } from 'react';
import { Link } from 'react-router-dom';
import { EnvelopeIcon, MapPinIcon, PhoneIcon } from '@heroicons/react/24/outline'; // Using Heroicons
import { motion } from 'framer-motion'; // For animations

const Contact = () => {
  const [name, setName] = useState('');
  const [email, setEmail] = useState('');
  const [message, setMessage] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [submissionSuccess, setSubmissionSuccess] = useState(false);
  const [submissionError, setSubmissionError] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    setIsSubmitting(true);
    setSubmissionSuccess(false);
    setSubmissionError('');

    // Simulate form submission (replace with your actual API call)
    setTimeout(() => {
      if (name && email && message) {
        console.log('Form submitted:', { name, email, message });
        setSubmissionSuccess(true);
        setName('');
        setEmail('');
        setMessage('');
      } else {
        setSubmissionError('Please fill in all fields.');
      }
      setIsSubmitting(false);
    }, 1500);
  };

  const containerVariants = {
    hidden: { opacity: 0 },
    visible: { opacity: 1, transition: { duration: 0.6, ease: 'easeInOut' } },
  };

  const infoVariants = {
    hidden: { opacity: 0, y: 20 },
    visible: { opacity: 1, y: 0, transition: { duration: 0.4, delay: 0.2, ease: 'easeOut' } },
  };

  const formVariants = {
    hidden: { opacity: 0, y: 30 },
    visible: { opacity: 1, y: 0, transition: { duration: 0.5, delay: 0.4, ease: 'easeOut' } },
  };

  const successMessageVariants = {
    hidden: { opacity: 0, scale: 0.8 },
    visible: { opacity: 1, scale: 1, transition: { duration: 0.3, delay: 0.6, ease: 'easeOut' } },
  };

  const errorMessageVariants = {
    hidden: { opacity: 0, y: -10 },
    visible: { opacity: 1, y: 0, transition: { duration: 0.3, delay: 0.6, ease: 'easeOut' } },
  };

  return (
    <motion.div
      className="bg-gradient-to-br from-gray-100 to-gray-200  py-12 px-4 sm:px-6 lg:px-8"
      variants={containerVariants}
      initial="hidden"
      animate="visible"
    >
      <div className="max-w-3xl mx-auto bg-white rounded-xl shadow-lg overflow-hidden">
        <div className="grid grid-cols-1 lg:grid-cols-2">
          {/* Contact Information Section */}
          <motion.div
            className="p-8 bg-indigo-600 text-white flex flex-col justify-center items-center text-center"
            variants={infoVariants}
            initial="hidden"
            animate="visible"
          >
            <h2 className="text-3xl font-bold mb-6">Get in Touch</h2>
            <p className="text-lg text-indigo-200 mb-8">
              We'd love to hear from you! Whether you have questions, feedback, or just want to say hello, feel free to reach out.
            </p>
            <div className="space-y-4">
              <div className="flex items-center justify-center gap-4">
                <EnvelopeIcon className="h-6 w-6 text-yellow-300" />
                <p className="text-indigo-100">
                  Email: <a href="mailto:support@easysheets.com" className="font-semibold hover:underline">support@easysheets.com</a>
                </p>
              </div>
              <div className="flex items-center justify-center gap-4">
                <MapPinIcon className="h-6 w-6 text-yellow-300" />
                <p className="text-indigo-100">
                  Address: Basistha, Ghy-29, Assam, Ind
                </p>
              </div>
              <div className="flex items-center justify-center gap-4">
                <PhoneIcon className="h-6 w-6 text-yellow-300" />
                <p className="text-indigo-100">
                  Phone: <a href="tel:+[Your Phone Number]" className="font-semibold hover:underline">+91 9876543210</a>
                </p>
              </div>
            </div>
            <div className="mt-8">
              <p className="text-sm text-indigo-300">Follow us on:</p>
              {/* Add social media icons/links here */}
              <div className="flex justify-center gap-4 mt-2">
                <Link to="#" className="text-indigo-200 hover:text-yellow-300 transition-colors">
                  {/* Social Icon */}
                  <svg className="w-6 h-6 fill-current" viewBox="0 0 24 24"><path fill="currentColor" d="M22.46 6c-.77.67-1.7.99-2.76 1.16 1.66-1 2.96-2.58 3.59-4.43-.8 1.43-2.01 2.48-3.47 3.05-1.46-1.55-3.54-2.53-5.88-2.53-4.64 0-8.4 3.76-8.4 8.4 0 .66.07 1.3.19 1.9-7-1.8-13.2-3.5-17.4-1.8 2.1 3.6 5.3 5.9 9.1 6.1-1.4-.4-2.8-.8-4.3-1.3.8 2.5 3.1 4.3 5.8 4.5-1.4.3-2.9.4-4.5.1 1.3 4 5.1 6.9 9.5 7.1-3.4 2.7-7.7 4.3-12.4 4.3-2.9 0-5.7-.8-8.1-2.3 4.5 2.9 9.8 4.6 15.7 4.6 18.9 0 29.2-15.6 29.2-29.2 0-.4 0-.8-.1-1.2 2-1.4 3.7-3.1 5-5.1z"/></svg>
                </Link>
                <Link to="#" className="text-indigo-200 hover:text-yellow-300 transition-colors">
                  {/* Social Icon */}
                  <svg className="w-6 h-6 fill-current" viewBox="0 0 24 24"><path fill="currentColor" d="M22 12c0 5.52-4.48 10-10 10S2 17.52 2 12 6.48 2 12 2s10 4.48 10 10zm-10-8c4.42 0 8 3.58 8 8s-3.58 8-8 8-8-3.58-8-8 3.58-8 8-8zm0 2c-3.31 0-6 2.69-6 6s2.69 6 6 6 6-2.69 6-6-2.69-6-6-6zm5 3c0 1.1-.9 2-2 2s-2-.9-2-2 .9-2 2-2 2 .9 2 2z"/></svg>
                </Link>
                {/* Add more social links */}
              </div>
            </div>
          </motion.div>

          {/* Contact Form Section */}
          <motion.div className="p-8" variants={formVariants} initial="hidden" animate="visible">
            <h2 className="text-2xl font-semibold text-gray-800 mb-6">Send Us a Message</h2>
            {submissionSuccess && (
              <motion.div
                className="bg-green-100 border border-green-400 text-green-700 px-4 py-3 rounded relative mb-4"
                role="alert"
                variants={successMessageVariants}
                initial="hidden"
                animate="visible"
              >
                <strong className="font-bold">Success!</strong>
                <span className="block sm:inline"> Your message has been sent. We'll get back to you soon.</span>
              </motion.div>
            )}
            {submissionError && (
              <motion.div
                className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded relative mb-4"
                role="alert"
                variants={errorMessageVariants}
                initial="hidden"
                animate="visible"
              >
                <strong className="font-bold">Error!</strong>
                <span className="block sm:inline">{submissionError}</span>
              </motion.div>
            )}
            <form onSubmit={handleSubmit} className="space-y-4">
              <div>
                <label htmlFor="name" className="block text-gray-700 text-sm font-bold mb-2">
                  Name
                </label>
                <input
                  type="text"
                  id="name"
                  className="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  placeholder="Your Name"
                />
              </div>
              <div>
                <label htmlFor="email" className="block text-gray-700 text-sm font-bold mb-2">
                  Email Address
                </label>
                <input
                  type="email"
                  id="email"
                  className="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  placeholder="Your Email"
                />
              </div>
              <div>
                <label htmlFor="message" className="block text-gray-700 text-sm font-bold mb-2">
                  Message
                </label>
                <textarea
                  id="message"
                  rows="4"
                  className="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                  value={message}
                  onChange={(e) => setMessage(e.target.value)}
                  placeholder="Your Message"
                ></textarea>
              </div>
              <button
                type="submit"
                className={`bg-indigo-500 hover:bg-indigo-700 text-white font-bold py-3 px-6 rounded focus:outline-none focus:shadow-outline transition-colors duration-300 ${
                  isSubmitting ? 'opacity-75 cursor-not-allowed' : ''
                }`}
                disabled={isSubmitting}
              >
                {isSubmitting ? 'Sending...' : 'Send Message'}
              </button>
            </form>
          </motion.div>
        </div>
      </div>
    </motion.div>
  );
};

export default Contact;