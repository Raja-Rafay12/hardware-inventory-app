import React, { useState } from 'react';
import { Loader2, KeyRound, Mail, AlertTriangle, Eye, EyeOff, Check, X, ShieldAlert } from 'lucide-react';

function Auth({ onAuthSuccess }) {
  const [mode, setMode] = useState('login'); // 'login', 'signup', 'forgot', 'otp'
  const [firstName, setFirstName] = useState('');
  const [lastName, setLastName] = useState('');
  const [organizationName, setOrganizationName] = useState('');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [confirmPassword, setConfirmPassword] = useState('');
  const [otp, setOtp] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [infoMessage, setInfoMessage] = useState(null);

  // Live Validations
  const isEmailValid = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  const hasMinLength = password.length >= 8;
  const hasCapital = /[A-Z]/.test(password);
  const hasNumber = /[0-9]/.test(password);
  const passwordsMatch = password === confirmPassword;

  let isFormValid = false;
  if (mode === 'login') {
    isFormValid = !!(email && password);
  } else if (mode === 'signup') {
    isFormValid = !!(
      firstName.trim() &&
      lastName.trim() &&
      organizationName.trim() &&
      isEmailValid &&
      hasMinLength &&
      hasCapital &&
      hasNumber &&
      passwordsMatch
    );
  } else if (mode === 'forgot') {
    isFormValid = isEmailValid;
  } else if (mode === 'otp') {
    isFormValid = otp.trim().length === 6 && hasMinLength && hasCapital && hasNumber && passwordsMatch;
  }

  const isSignUp = mode === 'signup';

  const handleAuth = async (e) => {
    e.preventDefault();
    setError(null);
    setInfoMessage(null);

    if (mode === 'signup') {
      if (!firstName.trim() || !lastName.trim() || !organizationName.trim()) {
        setError('Please fill in all profile fields');
        return;
      }
      if (!isEmailValid) {
        setError('Please enter a valid email address');
        return;
      }
      if (!hasMinLength || !hasCapital || !hasNumber) {
        setError('Password does not meet all security requirements');
        return;
      }
      if (!passwordsMatch) {
        setError('Passwords do not match');
        return;
      }
    } else if (mode === 'login') {
      if (!email || !password) {
        setError('Please fill in all fields');
        return;
      }
    } else if (mode === 'forgot') {
      if (!isEmailValid) {
        setError('Please enter a valid email address');
        return;
      }
    } else if (mode === 'otp') {
      if (otp.trim().length !== 6) {
        setError('Please enter the 6-digit code');
        return;
      }
      if (!hasMinLength || !hasCapital || !hasNumber) {
        setError('Password does not meet all security requirements');
        return;
      }
      if (!passwordsMatch) {
        setError('Passwords do not match');
        return;
      }
    }

    setLoading(true);

    try {
      if (mode === 'signup') {
        const user = await window.db.signup({
          firstName,
          lastName,
          organizationName,
          email,
          password
        });
        setInfoMessage('Account created successfully! Welcome email sent.');
        if (onAuthSuccess) {
          setTimeout(() => onAuthSuccess(user), 800);
        }
      } else if (mode === 'login') {
        const user = await window.db.login(email, password);
        if (onAuthSuccess) {
          onAuthSuccess(user);
        }
      } else if (mode === 'forgot') {
        await window.db.requestResetOtp(email);
        setInfoMessage('Verification code sent to your email. Please check your inbox.');
        // Clear password fields for reset
        setPassword('');
        setConfirmPassword('');
        setMode('otp');
      } else if (mode === 'otp') {
        await window.db.confirmPasswordReset(email, otp, password);
        setInfoMessage('Password reset successfully! Redirecting to login...');
        setOtp('');
        setPassword('');
        setConfirmPassword('');
        setTimeout(() => {
          setMode('login');
          setInfoMessage(null);
        }, 1500);
      }
    } catch (err) {
      console.error(err);
      setError(err.message || 'An error occurred during authentication');
    } finally {
      setLoading(false);
    }
  };

  const overlayMessage = 
    mode === 'signup' ? "Creating your account & sending welcome mail..." : 
    mode === 'login' ? "Logging in, please wait..." : 
    mode === 'forgot' ? "Sending verification code..." : "Resetting password...";

  return (
    <div style={styles.container}>
      {loading && <LoadingOverlay message={overlayMessage} />}
      <style>{`
        .auth-card {
          width: 100%;
          max-width: 440px;
          background: #FFFFFF;
          border: 1px solid #E0D9C9;
          border-radius: 12px;
          padding: 32px 28px;
          box-shadow: 0 4px 20px rgba(38, 36, 32, 0.05);
          font-family: 'Inter', sans-serif;
          color: #262420;
          box-sizing: border-box;
        }
        .auth-brand {
          display: flex;
          flex-direction: column;
          align-items: center;
          margin-bottom: 28px;
        }
        .auth-brand-mark {
          font-size: 24px;
          width: 48px;
          height: 48px;
          border-radius: 10px;
          background: #D9720B;
          color: #1C1A16;
          display: flex;
          align-items: center;
          justify-content: center;
          margin-bottom: 12px;
        }
        .auth-brand-name {
          font-family: 'Barlow Condensed', sans-serif;
          font-weight: 700;
          font-size: 24px;
          letter-spacing: 0.2px;
          text-transform: uppercase;
          line-height: 1.1;
        }
        .auth-brand-sub {
          font-size: 13px;
          color: #746C5E;
          margin-top: 4px;
        }
        .auth-field {
          margin-bottom: 18px;
        }
        .auth-label {
          display: block;
          font-size: 11.5px;
          font-weight: 700;
          color: #746C5E;
          margin-bottom: 6px;
          letter-spacing: 0.5px;
        }
        .auth-input-wrapper {
          position: relative;
          display: flex;
          align-items: center;
        }
        .auth-input-icon {
          position: absolute;
          left: 12px;
          color: #746C5E;
          pointer-events: none;
        }
        .auth-input {
          width: 100%;
          padding: 10px 12px 10px 38px;
          border: 1px solid #E0D9C9;
          border-radius: 8px;
          font-size: 14px;
          font-family: 'Inter', sans-serif;
          background: #F6F3EC;
          color: #262420;
          transition: border-color 0.2s, box-shadow 0.2s;
          outline: none;
          box-sizing: border-box;
        }
        .auth-input-profile {
          width: 100%;
          padding: 10px 12px;
          border: 1px solid #E0D9C9;
          border-radius: 8px;
          font-size: 14px;
          font-family: 'Inter', sans-serif;
          background: #F6F3EC;
          color: #262420;
          transition: border-color 0.2s, box-shadow 0.2s;
          outline: none;
          box-sizing: border-box;
        }
        .auth-input:focus, .auth-input-profile:focus {
          border-color: #D9720B;
          box-shadow: 0 0 0 3px rgba(217, 114, 11, 0.15);
        }
        .auth-eye-btn {
          position: absolute;
          right: 12px;
          border: none;
          background: transparent;
          color: #746C5E;
          cursor: pointer;
          padding: 0;
          display: flex;
          align-items: center;
        }
        .auth-eye-btn:hover {
          color: #262420;
        }
        .auth-btn {
          width: 100%;
          padding: 11px;
          border: none;
          border-radius: 8px;
          background: #D9720B;
          color: #FFFFFF;
          font-weight: 600;
          font-size: 14.5px;
          cursor: pointer;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          font-family: 'Inter', sans-serif;
          transition: background-color 0.2s;
        }
        .auth-btn:hover:not(:disabled) {
          background: #B85C05;
        }
        .auth-btn:disabled {
          opacity: 0.6;
          cursor: not-allowed;
        }
        .auth-alert {
          display: flex;
          align-items: flex-start;
          gap: 10px;
          padding: 12px;
          border-radius: 8px;
          font-size: 13px;
          line-height: 1.4;
          margin-bottom: 20px;
        }
        .auth-alert-error {
          background: #F6E3E1;
          color: #B33A3A;
          border: 1px solid rgba(179, 58, 58, 0.25);
        }
        .auth-alert-info {
          background: #E3EFE4;
          color: #3F7D4F;
          border: 1px solid rgba(63, 125, 79, 0.25);
        }
        .auth-toggle {
          text-align: center;
          margin-top: 20px;
          font-size: 13px;
          color: #746C5E;
        }
        .auth-toggle-btn {
          border: none;
          background: transparent;
          color: #D9720B;
          font-weight: 600;
          cursor: pointer;
          padding: 0;
          margin-left: 4px;
          font-family: 'Inter', sans-serif;
        }
        .auth-toggle-btn:hover {
          text-decoration: underline;
        }
        .validation-card {
          background: #FAF9F6;
          border: 1px solid #E5DFD3;
          border-radius: 8px;
          padding: 12px 14px;
          margin-bottom: 18px;
          font-size: 12px;
        }
        .validation-header {
          font-weight: 600;
          color: #746C5E;
          margin-bottom: 8px;
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .validation-item {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-bottom: 4px;
          color: #8C8476;
          transition: color 0.2s;
        }
        .validation-item.valid {
          color: #2E6A3E;
        }
        .validation-item.invalid {
          color: #A03333;
        }
      `}</style>
      <div className="auth-card">
        <div className="auth-brand">
          <div className="auth-brand-mark">⛏</div>
          <div className="auth-brand-name">Hardware Inventory</div>
          <div className="auth-brand-sub">Production Portal</div>
        </div>

        {error && (
          <div className="auth-alert auth-alert-error">
            <AlertTriangle size={18} style={{ flexShrink: 0, marginTop: 1 }} />
            <div>{error}</div>
          </div>
        )}

        {infoMessage && (
          <div className="auth-alert auth-alert-info">
            <div>{infoMessage}</div>
          </div>
        )}        <form onSubmit={handleAuth}>
          {/* signup fields */}
          {mode === 'signup' && (
            <>
              <div style={{ display: 'flex', gap: '12px', marginBottom: '18px' }}>
                <div style={{ flex: 1 }}>
                  <label className="auth-label">FIRST NAME</label>
                  <input
                    type="text"
                    className="auth-input-profile"
                    placeholder="John"
                    value={firstName}
                    onChange={(e) => setFirstName(e.target.value)}
                    disabled={loading}
                    required
                  />
                </div>
                <div style={{ flex: 1 }}>
                  <label className="auth-label">LAST NAME</label>
                  <input
                    type="text"
                    className="auth-input-profile"
                    placeholder="Doe"
                    value={lastName}
                    onChange={(e) => setLastName(e.target.value)}
                    disabled={loading}
                    required
                  />
                </div>
              </div>

              <div className="auth-field">
                <label className="auth-label">ORGANIZATION NAME</label>
                <input
                  type="text"
                  className="auth-input-profile"
                  placeholder="Acme Hardware Ltd"
                  value={organizationName}
                  onChange={(e) => setOrganizationName(e.target.value)}
                  disabled={loading}
                  required
                />
              </div>
            </>
          )}

          {/* email field - rendered in all modes except OTP */}
          {mode !== 'otp' && (
            <div className="auth-field">
              <label className="auth-label">EMAIL ADDRESS</label>
              <div className="auth-input-wrapper">
                <Mail size={16} className="auth-input-icon" />
                <input
                  type="email"
                  className="auth-input"
                  placeholder="your@email.com"
                  value={email}
                  onChange={(e) => setEmail(e.target.value)}
                  disabled={loading}
                  required
                />
              </div>
            </div>
          )}

          {/* OTP field - only in otp mode */}
          {mode === 'otp' && (
            <div className="auth-field">
              <label className="auth-label">6-DIGIT VERIFICATION CODE</label>
              <div className="auth-input-wrapper">
                <KeyRound size={16} className="auth-input-icon" />
                <input
                  type="text"
                  className="auth-input"
                  placeholder="123456"
                  maxLength={6}
                  value={otp}
                  onChange={(e) => setOtp(e.target.value.replace(/\D/g, ''))} // only digits
                  disabled={loading}
                  style={{ textAlign: 'center', letterSpacing: '4px', fontWeight: 'bold', fontSize: '16px' }}
                  required
                />
              </div>
            </div>
          )}

          {/* password field - rendered in signup, login, and otp modes */}
          {(mode === 'login' || mode === 'signup' || mode === 'otp') && (
            <div className="auth-field">
              <label className="auth-label">
                {mode === 'otp' ? 'NEW PASSWORD' : 'PASSWORD'}
              </label>
              <div className="auth-input-wrapper">
                <KeyRound size={16} className="auth-input-icon" />
                <input
                  type={showPassword ? 'text' : 'password'}
                  className="auth-input"
                  placeholder="••••••••"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  disabled={loading}
                  required
                />
                <button
                  type="button"
                  className="auth-eye-btn"
                  onClick={() => setShowPassword(!showPassword)}
                  disabled={loading}
                >
                  {showPassword ? <EyeOff size={16} /> : <Eye size={16} />}
                </button>
              </div>
              {/* forgot password link - only in login mode */}
              {mode === 'login' && (
                <div style={{ textAlign: 'right', marginTop: '6px' }}>
                  <button
                    type="button"
                    style={{ border: 'none', background: 'transparent', color: '#D9720B', fontSize: '12px', cursor: 'pointer', padding: 0 }}
                    onClick={() => {
                      setMode('forgot');
                      setError(null);
                      setInfoMessage(null);
                    }}
                  >
                    Forgot Password?
                  </button>
                </div>
              )}
            </div>
          )}

          {/* re-enter password field - in signup and otp modes */}
          {(mode === 'signup' || mode === 'otp') && (
            <>
              <div className="auth-field">
                <label className="auth-label">
                  {mode === 'otp' ? 'RE-ENTER NEW PASSWORD' : 'RE-ENTER PASSWORD'}
                </label>
                <div className="auth-input-wrapper">
                  <KeyRound size={16} className="auth-input-icon" />
                  <input
                    type="password"
                    className="auth-input"
                    placeholder="••••••••"
                    value={confirmPassword}
                    onChange={(e) => setConfirmPassword(e.target.value)}
                    disabled={loading}
                    required
                  />
                </div>
              </div>

              <div className="validation-card">
                <div className="validation-header">
                  <ShieldAlert size={14} style={{ color: '#D9720B' }} />
                  <span>Security Requirements</span>
                </div>
                {mode === 'signup' && (
                  <div className={`validation-item ${isEmailValid ? 'valid' : 'invalid'}`}>
                    {isEmailValid ? <Check size={13} style={{ color: '#2E6A3E' }} /> : <X size={13} style={{ color: '#A03333' }} />}
                    <span>Valid Email address (contains @)</span>
                  </div>
                )}
                <div className={`validation-item ${hasMinLength ? 'valid' : 'invalid'}`}>
                  {hasMinLength ? <Check size={13} style={{ color: '#2E6A3E' }} /> : <X size={13} style={{ color: '#A03333' }} />}
                  <span>Minimum 8 characters</span>
                </div>
                <div className={`validation-item ${hasCapital ? 'valid' : 'invalid'}`}>
                  {hasCapital ? <Check size={13} style={{ color: '#2E6A3E' }} /> : <X size={13} style={{ color: '#A03333' }} />}
                  <span>At least one capital letter (A-Z)</span>
                </div>
                <div className={`validation-item ${hasNumber ? 'valid' : 'invalid'}`}>
                  {hasNumber ? <Check size={13} style={{ color: '#2E6A3E' }} /> : <X size={13} style={{ color: '#A03333' }} />}
                  <span>At least one number (0-9)</span>
                </div>
                <div className={`validation-item ${passwordsMatch && password ? 'valid' : 'invalid'}`}>
                  {passwordsMatch && password ? <Check size={13} style={{ color: '#2E6A3E' }} /> : <X size={13} style={{ color: '#A03333' }} />}
                  <span>Passwords match</span>
                </div>
              </div>
            </>
          )}

          <button type="submit" className="auth-btn" disabled={loading || !isFormValid}>
            {loading ? (
              <>
                <Loader2 size={16} style={{ animation: 'spin 1s linear infinite' }} />
                {mode === 'signup' ? 'Creating Account...' :
                 mode === 'login' ? 'Logging In...' :
                 mode === 'forgot' ? 'Sending Code...' : 'Resetting Password...'}
              </>
            ) : (
              mode === 'signup' ? 'Create Account' :
              mode === 'login' ? 'Log In' :
              mode === 'forgot' ? 'Send Verification Code' : 'Reset Password'
            )}
          </button>
        </form>

        <div className="auth-toggle">
          {mode === 'signup' && (
            <>
              Already have an account?
              <button
                type="button"
                className="auth-toggle-btn"
                onClick={() => {
                  setMode('login');
                  setError(null);
                  setInfoMessage(null);
                  setFirstName('');
                  setLastName('');
                  setOrganizationName('');
                  setConfirmPassword('');
                }}
                disabled={loading}
              >
                Log In
              </button>
            </>
          )}
          {mode === 'login' && (
            <>
              Don't have an account?
              <button
                type="button"
                className="auth-toggle-btn"
                onClick={() => {
                  setMode('signup');
                  setError(null);
                  setInfoMessage(null);
                }}
                disabled={loading}
              >
                Create Account
              </button>
            </>
          )}
          {(mode === 'forgot' || mode === 'otp') && (
            <button
              type="button"
              className="auth-toggle-btn"
              onClick={() => {
                setMode('login');
                setError(null);
                setInfoMessage(null);
                setOtp('');
                setPassword('');
                setConfirmPassword('');
              }}
              disabled={loading}
              style={{ marginLeft: 0 }}
            >
              Back to Log In
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

const styles = {
  container: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: '100%',
    height: '100vh',
    background: '#F6F3EC',
    margin: 0,
    padding: 20,
    boxSizing: 'border-box',
  },
};

export default Auth;

function LoadingOverlay({ message = "Loading..." }) {
  return (
    <div style={{
      position: "fixed",
      top: 0, left: 0, right: 0, bottom: 0,
      background: "rgba(38, 36, 32, 0.7)",
      backdropFilter: "blur(4px)",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      zIndex: 9999,
      fontFamily: "'Inter', sans-serif",
      color: "#FFFFFF"
    }}>
      <div style={{
        background: "#262420",
        border: "1px solid #D9720B",
        borderRadius: "16px",
        padding: "32px 48px",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        boxShadow: "0 10px 40px rgba(0, 0, 0, 0.4)",
        animation: "authScaleIn 0.3s cubic-bezier(0.34, 1.56, 0.64, 1)"
      }}>
        <style>{`
          @keyframes authScaleIn {
            from { transform: scale(0.9); opacity: 0; }
            to { transform: scale(1); opacity: 1; }
          }
          @keyframes authSpinPulse {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .auth-loading-spinner {
            width: 52px;
            height: 52px;
            border: 4px solid rgba(217, 114, 11, 0.15);
            border-top-color: #D9720B;
            border-radius: 50%;
            animation: authSpinPulse 0.9s cubic-bezier(0.55, 0.085, 0.68, 0.53) infinite;
            margin-bottom: 20px;
            box-shadow: 0 0 15px rgba(217, 114, 11, 0.2);
          }
        `}</style>
        <div className="auth-loading-spinner"></div>
        <div style={{ fontSize: "16px", fontWeight: "600", letterSpacing: "0.2px", color: "#F6F3EC" }}>{message}</div>
      </div>
    </div>
  );
}
