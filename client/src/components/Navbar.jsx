import React, { useState, useRef, useEffect } from "react";
import { Link } from "react-router-dom";
import { useAuth } from "../hooks/useAuth";
import { motion, AnimatePresence } from "framer-motion"; // For more advanced animations
import {
  ShieldCheckIcon,
  ScaleIcon,
  InformationCircleIcon,
  ChatBubbleLeftRightIcon,
  ArrowRightOnRectangleIcon,
} from "@heroicons/react/24/outline"; // More relevant icons

const Navbar = () => {
  const [isOpen, setIsOpen] = useState(false);
  const dropdownRef = useRef(null);
  const { user, logout } = useAuth();

  useEffect(() => {
    const handleClickOutside = (event) => {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsOpen(false);
      }
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const closeDropdown = () => {
    setIsOpen(false);
  };

  const dropdownVariants = {
    hidden: {
      opacity: 0,
      y: -10,
      transition: { duration: 0.15, ease: "easeOut" },
    },
    visible: {
      opacity: 1,
      y: 0,
      transition: { duration: 0.2, ease: "easeOut" },
    },
    exit: {
      opacity: 0,
      y: -10,
      transition: { duration: 0.1, ease: "easeIn" },
    },
  };

  const menuItemVariants = {
    initial: { opacity: 0, y: -5 },
    animate: { opacity: 1, y: 0 },
    hover: {
      backgroundColor: "rgba(0, 0, 0, 0.05)",
      scale: 1.02,
      transition: { duration: 0.1 },
    },
    tap: { scale: 0.98 },
  };

  const logoutButtonVariants = {
    hover: { backgroundColor: "#dc2626" }, // Darker red on hover
    tap: { scale: 0.95 },
  };

  return (
    <nav className="bg-gradient-to-r from-indigo-600 to-purple-600 shadow-lg py-3 px-6 w-full z-50 backdrop-blur-md">
      <div className="container mx-auto flex items-center justify-between">
        {/* Logo */}
        <Link to="/" className="text-3xl font-extrabold text-white tracking-wider">
          EASY <span className="text-yellow-300">SHEETS</span>
        </Link>

        {/* Profile Menu */}
        <div className="relative" ref={dropdownRef}>
          <button
            onClick={() => setIsOpen(!isOpen)}
            className="focus:outline-none flex items-center gap-2"
            aria-label="User Menu"
          >
            <motion.img
              src={user?.profilePicture || "https://via.placeholder.com/40"}
              alt="Profile"
              className="w-12 h-12 rounded-full border-2 border-white hover:shadow-md transition-shadow"
              whileHover={{ scale: 1.05 }}
              whileTap={{ scale: 0.95 }}
            />
          </button>

          {/* Dropdown / Navigation Menu with AnimatePresence */}
          <AnimatePresence>
            {isOpen && (
              <motion.div
                className="absolute right-0 mt-3 w-64 bg-white text-gray-800 shadow-xl rounded-xl overflow-hidden"
                variants={dropdownVariants}
                initial="hidden"
                animate="visible"
                exit="exit"
              >
                <div className="px-5 py-4 border-b border-gray-200 bg-gray-50">
                  <motion.img
                    src={user?.profilePicture || "https://via.placeholder.com/40"}
                    alt="Profile"
                    className="w-16 h-16 rounded-full mx-auto border-2 border-indigo-500 mb-2"
                    initial={{ opacity: 0, scale: 0.8 }}
                    animate={{ opacity: 1, scale: 1 }}
                    transition={{ duration: 0.3, ease: "easeOut" }}
                  />
                  <p className="text-center font-semibold text-lg text-gray-900">{user?.name || "User Account"}</p>
                  <p className="text-center text-sm text-gray-500 italic">Logged in</p>
                </div>

                {/* Premium Nav Links */}
                <motion.ul className="text-gray-700 font-medium" initial="initial" animate="animate">
                  <motion.li
                    className="px-5 py-3 hover:bg-gray-100 transition-colors cursor-pointer flex items-center gap-3"
                    variants={menuItemVariants}
                    whileHover="hover"
                    whileTap="tap"
                    onClick={closeDropdown}
                  >
                    <ShieldCheckIcon className="h-5 w-5 text-indigo-500" />
                    <Link to="/privacy-policy" className="block w-full">
                      <span className="text-indigo-700">Privacy</span> & Security
                    </Link>
                  </motion.li>
                  <motion.li
                    className="px-5 py-3 hover:bg-gray-100 transition-colors cursor-pointer flex items-center gap-3"
                    variants={menuItemVariants}
                    whileHover="hover"
                    whileTap="tap"
                    onClick={closeDropdown}
                  >
                    <ScaleIcon className="h-5 w-5 text-blue-500" />
                    <Link to="/terms-of-service" className="block w-full">
                      Terms of <span className="text-blue-700">Service</span>
                    </Link>
                  </motion.li>
                  <motion.li
                    className="px-5 py-3 hover:bg-gray-100 transition-colors cursor-pointer flex items-center gap-3"
                    variants={menuItemVariants}
                    whileHover="hover"
                    whileTap="tap"
                    onClick={closeDropdown}
                  >
                    <InformationCircleIcon className="h-5 w-5 text-green-500" />
                    <Link to="/about" className="block w-full">
                      Learn <span className="text-green-700">About</span> Us
                    </Link>
                  </motion.li>
                  <motion.li
                    className="px-5 py-3 hover:bg-gray-100 transition-colors cursor-pointer flex items-center gap-3"
                    variants={menuItemVariants}
                    whileHover="hover"
                    whileTap="tap"
                    onClick={closeDropdown}
                  >
                    <ChatBubbleLeftRightIcon className="h-5 w-5 text-yellow-500" />
                    <Link to="/contact" className="block w-full">
                      Get in <span className="text-yellow-700">Touch</span>
                    </Link>
                  </motion.li>
                </motion.ul>

                <motion.div className="px-5 py-3 border-t border-gray-200" variants={menuItemVariants}>
                  <motion.button
                    onClick={() => {
                      logout();
                      closeDropdown();
                    }}
                    className="w-full bg-red-500 hover:bg-red-600 text-white py-2 rounded-md transition-colors font-semibold flex items-center justify-center gap-2"
                    variants={logoutButtonVariants}
                    whileHover="hover"
                    whileTap="tap"
                  >
                    <ArrowRightOnRectangleIcon className="h-5 w-5" />
                    Sign Out
                  </motion.button>
                </motion.div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>
    </nav>
  );
};

export default Navbar;