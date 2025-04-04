import React from 'react';
import { Link } from 'react-router-dom';
import { LightBulbIcon, PuzzlePieceIcon, RocketLaunchIcon, UserGroupIcon } from '@heroicons/react/24/outline';
import { motion } from 'framer-motion';
import { FaDribbble, FaLinkedin, FaTwitter, FaGithub } from 'react-icons/fa'; // Using react-icons for Dribbble

// Sample team member data (replace with your actual team)
const teamMembers = [
  {
    name: 'John Doe',
    title: 'Founder & CEO',
    image: 'https://via.placeholder.com/150/778899/fff?Text=JD',
    social: { linkedin: '#', twitter: '#' },
  },
  {
    name: 'Jane Smith',
    title: 'Lead Developer',
    image: 'https://via.placeholder.com/150/87CEEB/fff?Text=JS',
    social: { github: '#' },
  },
  {
    name: 'Peter Jones',
    title: 'UI/UX Designer',
    image: 'https://via.placeholder.com/150/ADD8E6/fff?Text=PJ',
    social: { dribbble: '#' },
  },
  // Add more team members
];

const About = () => {
  const containerVariants = {
    hidden: { opacity: 0 },
    visible: { opacity: 1, transition: { duration: 0.6, ease: 'easeInOut' } },
  };

  const sectionVariants = {
    hidden: { opacity: 0, y: 20 },
    visible: { opacity: 1, y: 0, transition: { duration: 0.5, delay: 0.2, ease: 'easeOut' } },
  };

  const teamMemberVariants = {
    hidden: { opacity: 0, scale: 0.8 },
    visible: { opacity: 1, scale: 1, transition: { duration: 0.4, delay: 0.3, ease: 'easeOut' } },
  };

  return (
    <motion.div
      className="bg-gradient-to-br from-gray-100 to-gray-200 min-h-screen py-12 px-4 sm:px-6 lg:px-8"
      variants={containerVariants}
      initial="hidden"
      animate="visible"
    >
      <div className="max-w-5xl mx-auto bg-white rounded-xl shadow-lg overflow-hidden">
        <div className="p-8">
          <motion.header className="text-center mb-10" variants={sectionVariants} initial="hidden" animate="visible">
            <h1 className="text-4xl font-extrabold text-indigo-700 mb-4">About Easy Sheets</h1>
            <p className="text-lg text-gray-600">
              Our mission is to simplify spreadsheet management and empower users with intuitive tools for data organization and collaboration.
            </p>
          </motion.header>

          <motion.section className="mb-10 grid grid-cols-1 md:grid-cols-2 gap-8" variants={sectionVariants} initial="hidden" animate="visible">
            <div>
              <LightBulbIcon className="h-12 w-12 text-yellow-500 mb-4" />
              <h2 className="text-2xl font-semibold text-gray-800 mb-3">Our Vision</h2>
              <p className="text-gray-700">
                We envision a world where working with data is effortless and accessible to everyone. Easy Sheets is designed to be the go-to platform for individuals and teams looking for a powerful yet user-friendly spreadsheet solution.
              </p>
            </div>
            <div>
              <PuzzlePieceIcon className="h-12 w-12 text-blue-500 mb-4" />
              <h2 className="text-2xl font-semibold text-gray-800 mb-3">Our Approach</h2>
              <p className="text-gray-700">
                We believe in continuous innovation and user-centric design. Our approach involves listening to our community, constantly improving our features, and ensuring a seamless and enjoyable user experience across all devices.
              </p>
            </div>
          </motion.section>

          <motion.section className="mb-10" variants={sectionVariants} initial="hidden" animate="visible">
            <h2 className="text-3xl font-semibold text-indigo-700 mb-6 text-center">Key Features</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              <div className="bg-gray-100 rounded-lg p-6 shadow-sm hover:shadow-md transition-shadow duration-300">
                <RocketLaunchIcon className="h-8 w-8 text-purple-500 mb-3" />
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Intuitive Interface</h3>
                <p className="text-gray-600">Designed for ease of use, even for those new to spreadsheets.</p>
              </div>
              <div className="bg-gray-100 rounded-lg p-6 shadow-sm hover:shadow-md transition-shadow duration-300">
                <UserGroupIcon className="h-8 w-8 text-green-500 mb-3" />
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Real-time Collaboration</h3>
                <p className="text-gray-600">Work together with your team seamlessly and see changes instantly.</p>
              </div>
              <div className="bg-gray-100 rounded-lg p-6 shadow-sm hover:shadow-md transition-shadow duration-300">
                {/* Add more feature icons */}
                <svg className="h-8 w-8 text-orange-500 mb-3" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 6h16M4 10h16M4 14h16M4 18h16" /></svg>
                <h3 className="text-xl font-semibold text-gray-800 mb-2">Versatile Functionality</h3>
                <p className="text-gray-600">From basic calculations to advanced data analysis tools.</p>
              </div>
              {/* Add more feature blocks */}
            </div>
          </motion.section>

          <motion.section className="mb-10" variants={sectionVariants} initial="hidden" animate="visible">
            <h2 className="text-3xl font-semibold text-indigo-700 mb-6 text-center">Our Team</h2>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-8">
              {teamMembers.map((member, index) => (
                <motion.div
                  key={index}
                  className="bg-white rounded-lg shadow-md p-6 text-center"
                  variants={teamMemberVariants}
                  initial="hidden"
                  animate="visible"
                >
                  <img
                    src={member.image}
                    alt={member.name}
                    className="w-24 h-24 rounded-full mx-auto mb-4 object-cover"
                  />
                  <h4 className="text-xl font-semibold text-gray-800 mb-2">{member.name}</h4>
                  <p className="text-gray-500 mb-3">{member.title}</p>
                  <div className="flex justify-center gap-4">
                    {member.social?.linkedin && (
                      <Link to={member.social.linkedin} className="text-indigo-500 hover:text-indigo-700">
                        <FaLinkedin className="w-5 h-5 fill-current" />
                      </Link>
                    )}
                    {member.social?.twitter && (
                      <Link to={member.social.twitter} className="text-blue-500 hover:text-blue-700">
                        <FaTwitter className="w-5 h-5 fill-current" />
                      </Link>
                    )}
                    {member.social?.github && (
                      <Link to={member.social.github} className="text-gray-800 hover:text-gray-900">
                        <FaGithub className="w-5 h-5 fill-current" />
                      </Link>
                    )}
                    {member.social?.dribbble && (
                      <Link to={member.social.dribbble} className="text-pink-500 hover:text-pink-700">
                        <FaDribbble className="w-5 h-5 fill-current" />
                      </Link>
                    )}
                    {/* Add more social icons as needed */}
                  </div>
                </motion.div>
              ))}
            </div>
          </motion.section>

          <motion.footer className="text-center text-gray-500 mt-12" variants={sectionVariants} initial="hidden" animate="visible">
            <p>&copy; {new Date().getFullYear()} Easy Sheets. All rights reserved.</p>
            <Link to="/privacy-policy" className="hover:underline mr-4">Privacy Policy</Link>
            <Link to="/terms-of-service" className="hover:underline">Terms of Service</Link>
          </motion.footer>
        </div>
      </div>
    </motion.div>
  );
};

export default About;