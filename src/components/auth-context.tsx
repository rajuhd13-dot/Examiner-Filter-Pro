import React, { createContext, useContext, useEffect, useState } from "react";
import { initializeApp } from "firebase/app";
import { 
  getAuth, 
  signInWithPopup, 
  GoogleAuthProvider, 
  onAuthStateChanged, 
  User, 
  UserCredential 
} from "firebase/auth";
import firebaseConfig from "../../firebase-applet-config.json";

// Shared scopes used by the app
export const SCOPES = [
  "https://www.googleapis.com/auth/spreadsheets.readonly"
];

interface AuthContextType {
  user: User | null;
  token: string | null;
  isLoggingIn: boolean;
  login: () => Promise<void>;
  logout: () => Promise<void>;
  initialized: boolean;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const AuthProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [user, setUser] = useState<User | null>(null);
  const [token, setToken] = useState<string | null>(null);
  const [isLoggingIn, setIsLoggingIn] = useState(false);
  const [initialized, setInitialized] = useState(false);

  // Initialize Firebase
  const app = initializeApp(firebaseConfig);
  const auth = getAuth(app);

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      // Note: Firebase doesn't always provide the OAuth token in onAuthStateChanged
      // We rely on the popup result or manual refresh logic if needed.
      // For this app, we'll store the token from login in state.
      setInitialized(true);
    });
    return () => unsubscribe();
  }, [auth]);

  const login = async () => {
    setIsLoggingIn(true);
    try {
      const provider = new GoogleAuthProvider();
      SCOPES.forEach(scope => provider.addScope(scope));
      
      const result: UserCredential = await signInWithPopup(auth, provider);
      const credential = GoogleAuthProvider.credentialFromResult(result);
      
      if (credential?.accessToken) {
        setToken(credential.accessToken);
        setUser(result.user);
      } else {
        console.error("No access token returned from Google login");
      }
    } catch (error) {
      console.error("Login failed:", error);
    } finally {
      setIsLoggingIn(false);
    }
  };

  const logout = async () => {
    await auth.signOut();
    setToken(null);
    setUser(null);
  };

  return (
    <AuthContext.Provider value={{ user, token, isLoggingIn, login, logout, initialized }}>
      {children}
    </AuthContext.Provider>
  );
};

export const useAuth = () => {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error("useAuth must be used within an AuthProvider");
  }
  return context;
};
