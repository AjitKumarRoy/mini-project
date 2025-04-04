import axios from 'axios'

const api = axios.create({
  baseURL: import.meta.env.VITE_API_BASE_URL || 'http://localhost:3000',
  withCredentials: true, // important if your backend sets HttpOnly cookies
})

// Optional: Add interceptors for error handling or request headers
api.interceptors.response.use(
  response => response,
  error => {
    // global error handling
    return Promise.reject(error)
  }
)

export default api
