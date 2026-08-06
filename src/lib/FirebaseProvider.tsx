import React, { createContext, useContext, useEffect, useState } from 'react';
import logoImg from '../assets/logo.png';
import { getSupabaseClient } from './supabase';
import { firestoreService } from './firestoreService';

export type UserRole = 'coordenador_geral' | 'coordenador_regional' | 'lider' | 'coordenador';

export interface UserProfile {
  uid: string;
  email: string | null;
  displayName?: string | null;
  name?: string | null;
  role?: UserRole;
  region?: string | null;
  coordinatorId?: string | null;
  teamId?: string | null;
  teamName?: string | null;
  forcePasswordChange?: boolean;
}

interface AuthContextType {
  user: any;
  role: UserRole | null;
  loading: boolean;
  forcePasswordChange: boolean;
  login: () => Promise<void>;
  loginWithEmail: (email: string, pass: string) => Promise<void>;
  signupWithEmail: (email: string, pass: string, role: UserRole, extraData?: any) => Promise<void>;
  logout: () => Promise<void>;
  changePassword: (newPass: string) => Promise<void>;
  resetPassword: (email: string) => Promise<void>;
  verifyEmail: () => Promise<void>;
  isAdmin: boolean;
  isGeral: boolean;
  isRegional: boolean;
  isLeader: boolean;
  userRegion: string | null;
  coordinatorId: string | null;
  demoRole: UserRole | null;
  setDemoRole: (role: UserRole | null) => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export function FirebaseProvider({ children }: { children: React.ReactNode }) {
  const [user, setUser] = useState<any>(null);
  const [role, setRole] = useState<UserRole | null>(null);
  const [demoRole, setDemoRole] = useState<UserRole | null>(null);
  const [loading, setLoading] = useState(true);
  const [isAdmin, setIsAdmin] = useState(false);
  const [isGeral, setIsGeral] = useState(false);
  const [isRegional, setIsRegional] = useState(false);
  const [isLeader, setIsLeader] = useState(false);
  const [userRegion, setUserRegion] = useState<string | null>(null);
  const [forcePasswordChange, setForcePasswordChange] = useState(false);
  const [coordinatorId, setCoordinatorId] = useState<string | null>(null);

  // Helper to sync user profile state
  const syncUserProfile = async (authUser: any) => {
    if (!authUser) {
      setUser(null);
      setRole(null);
      setIsAdmin(false);
      setIsGeral(false);
      setIsRegional(false);
      setIsLeader(false);
      setUserRegion(null);
      setCoordinatorId(null);
      setLoading(false);
      return;
    }

    // A real logged-in user overrides any active demo mode
    setDemoRole(null);

    const uid = authUser.id || authUser.uid;
    const cleanEmail = (authUser.email || '').toLowerCase().trim();
    const emailPrefix = cleanEmail.split('@')[0];
    const isSergioGeral = cleanEmail.includes('sergio') || cleanEmail.includes('sergiosilvabezerra');

    // Check URL parameters for explicit role hint
    const urlParams = new URLSearchParams(window.location.search);
    const urlRole = urlParams.get('role');
    const urlEmail = (urlParams.get('email') || '').toLowerCase().trim();

    // Fetch user profile from Supabase or local users collection
    let profile: any = await firestoreService.getDocument('users', uid);

    if (!profile && cleanEmail) {
      const allUsers = await firestoreService.getCollection<any>('users');
      profile = allUsers.find(u => {
        if (!u) return false;
        if (u.id === uid || u.uid === uid) return true;
        if (!u.email) return false;
        const uE = u.email.toLowerCase().trim();
        return uE === cleanEmail || uE.split('@')[0] === emailPrefix || (urlEmail && uE === urlEmail);
      });
      if (profile) {
        profile.uid = uid;
        profile.id = uid;
      }
    }

    // Search regional_coordinators and pre_registrations
    const allRegs = await firestoreService.getCollection<any>('regional_coordinators');
    const regCoord = allRegs.find(r => {
      if (!r) return false;
      const rEmail = (r.email || '').toLowerCase().trim();
      const rName = (r.name || '').toLowerCase().trim();
      return (
        (rEmail && (rEmail === cleanEmail || rEmail.split('@')[0] === emailPrefix || (urlEmail && rEmail === urlEmail))) ||
        (rName && cleanEmail && emailPrefix.length >= 3 && (rName.includes(emailPrefix) || emailPrefix.includes(rName)))
      );
    });

    const allPreRegs = await firestoreService.getCollection<any>('pre_registrations');
    const preReg = allPreRegs.find(p => {
      if (!p) return false;
      const pEmail = (p.email || '').toLowerCase().trim();
      const pName = (p.name || '').toLowerCase().trim();
      return (
        (pEmail && (pEmail === cleanEmail || pEmail.split('@')[0] === emailPrefix || (urlEmail && pEmail === urlEmail))) ||
        (pName && cleanEmail && (pName.includes(emailPrefix) || emailPrefix.includes(pName)))
      );
    }) || (cleanEmail ? await firestoreService.getDocument('pre_registrations', cleanEmail) : null);

    // Search teams collection for team leader match
    const allTeams = await firestoreService.getCollection<any>('teams');
    const teamLeaderDoc = allTeams.find(t => {
      if (!t) return false;
      const lEmail = (t.leaderEmail || '').toLowerCase().trim();
      const lName = (t.leader || '').toLowerCase().trim();
      return (
        (lEmail && (lEmail === cleanEmail || lEmail.split('@')[0] === emailPrefix || (urlEmail && lEmail === urlEmail))) ||
        (lName && cleanEmail && (
          (emailPrefix.length > 2 && (lName.includes(emailPrefix) || emailPrefix.includes(lName))) ||
          cleanEmail.includes(lName) || lName.includes(cleanEmail)
        ))
      );
    });

    // Determine leader status first
    const isLeaderUser = !isSergioGeral && (
      urlRole === 'lider' ||
      !!teamLeaderDoc ||
      (preReg && preReg.role === 'lider') ||
      (profile && profile.role === 'lider')
    );

    // Auto-cleanup stale regional_coordinator entry if user is a team leader
    if (isLeaderUser && regCoord && regCoord.id && (!preReg || preReg.role !== 'coordenador_regional')) {
      firestoreService.deleteDocument('regional_coordinators', regCoord.id).catch(() => {});
    }

    // Determine regional status
    const isRegionalUser = !isSergioGeral && !isLeaderUser && (
      (regCoord && regCoord.role !== 'lider') ||
      urlRole === 'coordenador_regional' ||
      (preReg && preReg.role === 'coordenador_regional') ||
      (profile && profile.role === 'coordenador_regional' && !teamLeaderDoc)
    );

    let currentRole: UserRole = 'coordenador_regional';

    if (isSergioGeral) {
      currentRole = 'coordenador_geral';
    } else if (isLeaderUser) {
      currentRole = 'lider';
    } else if (isRegionalUser) {
      currentRole = 'coordenador_regional';
    } else if (preReg && preReg.role) {
      currentRole = preReg.role;
    } else {
      currentRole = 'coordenador_regional';
    }

    const parentCoordId = regCoord?.coordinatorId || preReg?.coordinatorId || profile?.coordinatorId || teamLeaderDoc?.coordinatorId || (currentRole === 'coordenador_geral' ? uid : 'geral');
    const regRegion = regCoord?.region || preReg?.region || profile?.region || teamLeaderDoc?.location || 'REGIÃO SUL';
    const regName = preReg?.name || teamLeaderDoc?.leader || regCoord?.name || profile?.name || authUser.user_metadata?.full_name || authUser.displayName || (currentRole === 'lider' ? 'Líder de Equipe' : 'Coordenador Regional');

    profile = {
      ...(profile || {}),
      id: uid,
      uid,
      email: cleanEmail || regCoord?.email || preReg?.email || teamLeaderDoc?.leaderEmail,
      role: currentRole,
      name: regName,
      region: regRegion,
      coordinatorId: parentCoordId,
      teamId: teamLeaderDoc?.id || preReg?.teamId || profile?.teamId,
      teamName: teamLeaderDoc?.name || preReg?.teamName || profile?.teamName,
      regionalCoordId: teamLeaderDoc?.regionalCoordId || preReg?.regionalCoordId || profile?.regionalCoordId || '',
      regionalCoordEmail: teamLeaderDoc?.regionalCoordEmail || preReg?.regionalCoordEmail || profile?.regionalCoordEmail || '',
      createdAt: profile?.createdAt || Date.now()
    };

    // Save/update profile in Firestore
    await firestoreService.setDocument('users', uid, profile, true);
    if (cleanEmail) {
      await firestoreService.setDocument('users', `user_${cleanEmail}`, profile, true);
    }

    const regionalCheck = currentRole === 'coordenador_regional';
    const geralCheck = currentRole === 'coordenador_geral' && !regionalCheck;
    const leaderCheck = currentRole === 'lider';
    const adminCheck = geralCheck || regionalCheck;

    setUser({
      uid,
      id: uid,
      email: cleanEmail,
      displayName: regName,
      emailVerified: true
    });
    setRole(currentRole);
    setIsGeral(geralCheck);
    setIsRegional(regionalCheck);
    setIsLeader(leaderCheck);
    setIsAdmin(adminCheck);
    setUserRegion(profile.region || (regionalCheck ? 'REGIÃO SUL' : null));
    setForcePasswordChange(!!profile.forcePasswordChange);
    setCoordinatorId(geralCheck ? 'geral' : (parentCoordId || 'geral'));
    setLoading(false);
  };

  useEffect(() => {
    const supabase = getSupabaseClient();
    if (supabase) {
      // Fetch initial session
      supabase.auth.getSession().then(({ data: { session } }) => {
        if (session?.user) {
          syncUserProfile(session.user);
        } else {
          // Check local cached session fallback
          const localUser = localStorage.getItem('nexus_auth_user');
          if (localUser) {
            try {
              syncUserProfile(JSON.parse(localUser));
            } catch (e) {
              setLoading(false);
            }
          } else {
            setLoading(false);
          }
        }
      }).catch(() => setLoading(false));

      // Listen to auth changes
      const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
        if (session?.user) {
          localStorage.setItem('nexus_auth_user', JSON.stringify(session.user));
          syncUserProfile(session.user);
        } else {
          localStorage.removeItem('nexus_auth_user');
          syncUserProfile(null);
        }
      });

      return () => {
        subscription.unsubscribe();
      };
    } else {
      // Local storage auth fallback if Supabase not yet connected
      const localUser = localStorage.getItem('nexus_auth_user');
      if (localUser) {
        try {
          syncUserProfile(JSON.parse(localUser));
        } catch (e) {
          setLoading(false);
        }
      } else {
        setLoading(false);
      }
    }
  }, []);

  const login = async () => {
    const supabase = getSupabaseClient();
    if (supabase) {
      const { error } = await supabase.auth.signInWithOAuth({
        provider: 'google',
        options: {
          redirectTo: window.location.origin
        }
      });
      if (error) throw error;
    } else {
      throw new Error('Supabase não configurado. Clique no botão "Configurar Supabase" abaixo para inserir sua URL e Anon Key.');
    }
  };

  const loginWithEmail = async (email: string, pass: string) => {
    const supabase = getSupabaseClient();
    if (supabase) {
      const { data, error } = await supabase.auth.signInWithPassword({
        email,
        password: pass
      });
      if (error) throw error;
      if (data.user) {
        localStorage.setItem('nexus_auth_user', JSON.stringify(data.user));
        await syncUserProfile(data.user);
      }
    } else {
      // Fallback local login for dev/offline mode
      const fakeUser = {
        id: `usr_${Date.now()}`,
        uid: `usr_${Date.now()}`,
        email,
        user_metadata: { full_name: email.split('@')[0] }
      };
      localStorage.setItem('nexus_auth_user', JSON.stringify(fakeUser));
      await syncUserProfile(fakeUser);
    }
  };

  const signupWithEmail = async (email: string, pass: string, userRole: UserRole, extraData?: any) => {
    const supabase = getSupabaseClient();
    let userId = `usr_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;

    if (supabase) {
      const { data, error } = await supabase.auth.signUp({
        email,
        password: pass
      });
      if (error) throw error;
      if (data.user) {
        userId = data.user.id;
      }
    }

    const profileData = {
      id: userId,
      uid: userId,
      email: email.toLowerCase(),
      role: userRole,
      createdAt: Date.now(),
      ...extraData
    };

    await firestoreService.setDocument('users', userId, profileData, true);
    localStorage.setItem('nexus_auth_user', JSON.stringify(profileData));
    await syncUserProfile(profileData);
  };

  const logout = async () => {
    const supabase = getSupabaseClient();
    if (supabase) {
      await supabase.auth.signOut().catch(() => {});
    }
    localStorage.removeItem('nexus_auth_user');
    await syncUserProfile(null);
  };

  const changePassword = async (newPass: string) => {
    const supabase = getSupabaseClient();
    if (supabase) {
      const { error } = await supabase.auth.updateUser({ password: newPass });
      if (error) throw error;
    }
    if (user?.uid) {
      await firestoreService.setDocument('users', user.uid, { forcePasswordChange: false }, true);
      setForcePasswordChange(false);
    }
  };

  const resetPassword = async (userEmail: string) => {
    const supabase = getSupabaseClient();
    if (supabase) {
      const { error } = await supabase.auth.resetPasswordForEmail(userEmail, {
        redirectTo: `${window.location.origin}/reset-password`
      });
      if (error) throw error;
    }
  };

  const verifyEmail = async () => {
    // Handled automatically by Supabase Auth
  };

  // Compute effective auth context values for Demo Mode or normal Auth
  let effectiveUser = user;
  let effectiveRole = role;
  let effectiveIsAdmin = isAdmin;
  let effectiveIsGeral = isGeral;
  let effectiveIsRegional = isRegional;
  let effectiveIsLeader = isLeader;
  let effectiveUserRegion = userRegion;
  let effectiveCoordinatorId = coordinatorId;

  if (demoRole) {
    if (demoRole === 'coordenador_geral') {
      effectiveUser = {
        uid: 'demo_coord_geral',
        email: 'geral@nexuspolitica.com.br',
        displayName: 'Coordenador Geral (Demo)',
        emailVerified: true
      };
      effectiveRole = 'coordenador_geral';
      effectiveIsGeral = true;
      effectiveIsRegional = false;
      effectiveIsLeader = false;
      effectiveIsAdmin = true;
      effectiveUserRegion = null;
      effectiveCoordinatorId = 'demo_coord_geral';
    } else if (demoRole === 'coordenador_regional') {
      effectiveUser = {
        uid: 'demo_coord_regional',
        email: 'regional.norte@nexuspolitica.com.br',
        displayName: 'Coordenador Regional (Demo)',
        emailVerified: true
      };
      effectiveRole = 'coordenador_regional';
      effectiveIsGeral = false;
      effectiveIsRegional = true;
      effectiveIsLeader = false;
      effectiveIsAdmin = true;
      effectiveUserRegion = 'REGIÃO 1 - NORTE';
      effectiveCoordinatorId = 'demo_coord_geral';
    } else if (demoRole === 'lider') {
      effectiveUser = {
        uid: 'demo_lider',
        email: 'lider.bairro@nexuspolitica.com.br',
        displayName: 'Líder de Bairro (Demo)',
        emailVerified: true
      };
      effectiveRole = 'lider';
      effectiveIsGeral = false;
      effectiveIsRegional = false;
      effectiveIsLeader = true;
      effectiveIsAdmin = false;
      effectiveUserRegion = 'REGIÃO 1 - NORTE';
      effectiveCoordinatorId = 'demo_coord_geral';
    }
  }

  return (
    <AuthContext.Provider value={{ 
      user: effectiveUser, 
      role: effectiveRole, 
      loading: demoRole ? false : loading, 
      login, 
      loginWithEmail, 
      signupWithEmail, 
      logout, 
      isAdmin: effectiveIsAdmin, 
      isGeral: effectiveIsGeral, 
      isRegional: effectiveIsRegional, 
      isLeader: effectiveIsLeader, 
      userRegion: effectiveUserRegion, 
      forcePasswordChange, 
      changePassword, 
      resetPassword, 
      verifyEmail,
      coordinatorId: effectiveCoordinatorId,
      demoRole,
      setDemoRole
    }}>
      {(!loading || effectiveUser) ? children : (
        <div className="min-h-screen bg-zinc-950 flex flex-col items-center justify-center p-8 select-none">
          <div className="relative flex flex-col items-center max-w-sm w-full text-center">
            <div className="relative mb-6">
              <div className="absolute -inset-2 bg-blue-600/20 rounded-full blur-xl animate-pulse" />
              <img 
                src={logoImg} 
                onError={(e) => { 
                  const t = e.currentTarget; 
                  if (!t.dataset.fallback) { 
                    t.dataset.fallback = 'true'; 
                    t.src = '/logo.png'; 
                  } 
                }} 
                alt="Logo Nexus Política" 
                className="relative w-24 h-24 object-contain mx-auto drop-shadow-md"
              />
            </div>
            
            <div className="w-10 h-10 border-3 border-blue-500/30 border-t-blue-500 rounded-full animate-spin mb-4"></div>
            
            <h2 className="text-xl font-bold text-white tracking-tight mb-1">Nexus Política</h2>
            <p className="text-xs text-zinc-400 font-medium uppercase tracking-widest">Conectado ao Supabase Cloud...</p>
          </div>
        </div>
      )}
    </AuthContext.Provider>
  );
}

export function useAuth() {
  const context = useContext(AuthContext);
  if (context === undefined) {
    throw new Error('useAuth must be used within a FirebaseProvider');
  }
  return context;
}
