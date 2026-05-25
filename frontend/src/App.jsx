import { useState, useEffect } from 'react'
import './App.css'

function App() {
  const [message, setMessage] = useState('')

  useEffect(() => {
    fetch('http://localhost:8000/api/hello')
      .then(response => response.json())
      .then(data => setMessage(data.message))
      .catch(error => console.error('Error fetching data:', error))
  }, [])

  return (
    <div style={{ textAlign: 'center', padding: '50px' }}>
      <h1>React + Python FastAPI</h1>
      <div className="card" style={{ padding: '20px', borderRadius: '8px', border: '1px solid #ccc', marginTop: '20px' }}>
        <p>Message from backend: <strong>{message || 'Loading...'}</strong></p>
      </div>
    </div>
  )
}

export default App
