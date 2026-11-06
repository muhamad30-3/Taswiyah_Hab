import { useState } from 'react';
import { LoginScreen } from './components/login-screen';
import { AttendanceKiosk } from './components/attendance-kiosk';
import { AppShell } from './components/app-shell';
import { GlobalLoadingProvider } from './components/global-loading';

export default function App() {
  return <GlobalLoadingProvider><AppContent /></GlobalLoadingProvider>;
}

function AppContent() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const attendanceToken = new URLSearchParams(window.location.search).get("attendance")

  if (attendanceToken) return <AttendanceKiosk token={attendanceToken} />
  if (!isAuthenticated) {
    return <LoginScreen onLogin={() => setIsAuthenticated(true)} />;
  }

  return <GlobalLoadingProvider><AppShell /></GlobalLoadingProvider>;
}