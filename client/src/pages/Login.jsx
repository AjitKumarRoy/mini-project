import React from 'react';
import { useAuth } from '../hooks/useAuth';
import { FcGoogle } from 'react-icons/fc'; // Import the Google icon
import api from "../api/api";

const Login = () => {
  const { googleLogin } = useAuth(); // Assuming you have a googleLogin function in your useAuth hook

  const handleGoogleLogin = () => {
    window.location.href = api.defaults.baseURL + '/api/auth/google'; // Use the baseURL from api.js
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center py-12 px-4 sm:px-6 lg:px-8">
      {/* Animated Background (Subtle Circles) */}
      <div className="absolute top-0 left-0 w-full h-full overflow-hidden pointer-events-none">
        <div className="absolute bg-white opacity-10 rounded-full animate-pulse" style={{ top: '10%', left: '15%', width: '60px', height: '60px' }}></div>
        <div className="absolute bg-white opacity-15 rounded-full animate-bounce" style={{ top: '30%', right: '20%', width: '80px', height: '80px', animationDelay: '0.5s' }}></div>
        <div className="absolute bg-white opacity-10 rounded-full animate-pulse" style={{ bottom: '15%', left: '35%', width: '40px', height: '40px', animationDelay: '1s' }}></div>
        <div className="absolute bg-white opacity-20 rounded-full animate-bounce" style={{ bottom: '25%', right: '10%', width: '100px', height: '100px', animationDelay: '1.5s' }}></div>
        <div className="absolute bg-white opacity-15 rounded-full animate-pulse" style={{ top: '45%', left: '50%', width: '70px', height: '70px', animationDelay: '2s' }}></div>
      </div>

      <div className="relative bg-white shadow-xl rounded-lg p-8 max-w-md w-full ">
        <h2 className="text-2xl font-bold text-gray-800 text-center mb-6 text-3xl font-extrabold text-black tracking-wider">EASY  <span className="text-yellow-300">SHEETS</span></h2>

        <button
          onClick={handleGoogleLogin}
          className="group relative w-full flex justify-center py-3 px-4 border border-transparent text-sm font-medium rounded-md text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-all duration-300"
        >
          <span className="absolute left-4 inset-y-0 flex items-center">
            <FcGoogle className="h-6 w-6" aria-hidden="true" />
          </span>
          Sign in with Google
        </button>
      </div>
    </div>
  );
};

export default Login;