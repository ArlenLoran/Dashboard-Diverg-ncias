import { useState, useEffect } from 'react';
import { 
  Plus, Edit2, Trash2, ChevronRight, Settings, Layout, 
  Activity, Shield, Clock, BookOpen, Database, Save, X,
  AlertTriangle, Filter, ArrowLeft, Lock, GripVertical, ChevronUp, ChevronDown, Mail, HelpCircle
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { Section, Metric } from '../types';
import { 
  fetchDashboardConfig, 
  ensureSharePointConfig,
  addDivision, 
  updateDivision, 
  deleteDivision, 
  addMetric, 
  updateMetric, 
  deleteMetric,
  isUserAllowed,
  fetchAllowedUsers,
  addAllowedUser,
  removeAllowedUser,
  saveDivisionsIndices,
  saveMetricsIndices,
  getTeamsChatId,
  saveTeamsChatId,
  getEmailAlertsEnabled,
  saveEmailAlertsEnabled,
  getTeamsAlertsEnabled,
  saveTeamsAlertsEnabled,
  fetchAlertEmails,
  addAlertEmail,
  removeAlertEmail
} from '../services/configService';
import { getCurrentSharePointUserEmail, hasSpContext } from '../services/spService';

export function Admin() {
  const [sections, setSections] = useState<Section[]>([]);
  const [isLoading, setIsLoading] = useState(true);
  const [isSaving, setIsSaving] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error', text: string } | null>(null);
  const [editingSection, setEditingSection] = useState<Section | null>(null);
  const [editingMetric, setEditingMetric] = useState<{ metric: Metric, divisionId: string } | null>(null);
  const [isSectionModalOpen, setIsSectionModalOpen] = useState(false);
  const [isMetricModalOpen, setIsMetricModalOpen] = useState(false);

  // Permission settings
  const [isCheckingAccess, setIsCheckingAccess] = useState(true);
  const [hasAccess, setHasAccess] = useState(false);
  const [userEmail, setUserEmail] = useState('');
  const [isAccessModalOpen, setIsAccessModalOpen] = useState(false);
  const [allowedUsers, setAllowedUsers] = useState<{ id: string; email: string }[]>([]);
  const [newAccessEmail, setNewAccessEmail] = useState('');
  const [isSavingAccess, setIsSavingAccess] = useState(false);
  const [accessMessage, setAccessMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  // Email Notification Settings
  const [emailAlertsEnabled, setEmailAlertsEnabled] = useState(false);
  const [isEmailModalOpen, setIsEmailModalOpen] = useState(false);
  const [alertEmails, setAlertEmails] = useState<{ id: string; email: string }[]>([]);
  const [newAlertEmail, setNewAlertEmail] = useState('');
  const [isSavingEmails, setIsSavingEmails] = useState(false);
  const [emailConfigMessage, setEmailConfigMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  // Teams Settings
  const [teamsChatId, setTeamsChatId] = useState('');
  const [isSavingTeamsChatId, setIsSavingTeamsChatId] = useState(false);
  const [teamsMessage, setTeamsMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [teamsAlertsEnabled, setTeamsAlertsEnabled] = useState(false);
  const [isHelpModalOpen, setIsHelpModalOpen] = useState(false);
  const [testTeamsLink, setTestTeamsLink] = useState('');
  const [extractedIdResult, setExtractedIdResult] = useState('');

  const [draggedSectionIndex, setDraggedSectionIndex] = useState<number | null>(null);
  const [draggedMetricIndex, setDraggedMetricIndex] = useState<{ sectionId: string, index: number } | null>(null);

  const handleMoveSection = async (fromIdx: number, toIdx: number) => {
    if (toIdx < 0 || toIdx >= sections.length) return;
    const updated = [...sections];
    const [removed] = updated.splice(fromIdx, 1);
    updated.splice(toIdx, 0, removed);
    setSections(updated);
    try {
      await saveDivisionsIndices(updated.map(s => s.id!));
    } catch (e) {
      console.error(e);
    }
  };

  const handleMoveMetric = async (sectionId: string, fromIdx: number, toIdx: number) => {
    const section = sections.find(s => s.id === sectionId);
    if (!section || toIdx < 0 || toIdx >= section.metrics.length) return;
    const updatedMetrics = [...section.metrics];
    const [removed] = updatedMetrics.splice(fromIdx, 1);
    updatedMetrics.splice(toIdx, 0, removed);
    
    const updatedSections = sections.map(s => {
      if (s.id === sectionId) {
        return { ...s, metrics: updatedMetrics };
      }
      return s;
    });
    setSections(updatedSections);
    try {
      await saveMetricsIndices(sectionId, updatedMetrics.map(m => m.id));
    } catch (e) {
      console.error(e);
    }
  };

  const loadAllowedUsers = async () => {
    try {
      const users = await fetchAllowedUsers();
      setAllowedUsers(users);
    } catch (err) {
      console.error(err);
    }
  };

  useEffect(() => {
    if (isAccessModalOpen) {
      loadAllowedUsers();
    }
  }, [isAccessModalOpen]);

  const handleAddAccess = async () => {
    if (!newAccessEmail.trim() || !newAccessEmail.includes('@')) {
      setAccessMessage({ type: 'error', text: 'Insira um e-mail válido' });
      return;
    }
    setIsSavingAccess(true);
    setAccessMessage(null);
    try {
      const ok = await addAllowedUser(newAccessEmail);
      if (ok) {
        setAccessMessage({ type: 'success', text: 'Acesso concedido com sucesso!' });
        setNewAccessEmail('');
        await loadAllowedUsers();
      } else {
        setAccessMessage({ type: 'error', text: 'Erro ao conceder acesso' });
      }
    } catch (err: any) {
      setAccessMessage({ type: 'error', text: err?.message || 'Erro ao conceder acesso' });
    } finally {
      setIsSavingAccess(false);
    }
  };

  const handleRemoveAccess = async (id: string, email: string) => {
    if (window.confirm(`Deseja realmente remover o acesso de ${email}?`)) {
      setIsSavingAccess(true);
      setAccessMessage(null);
      try {
        const ok = await removeAllowedUser(id, email);
        if (ok) {
          setAccessMessage({ type: 'success', text: 'Acesso removido com sucesso!' });
          await loadAllowedUsers();
          
          const curEmail = getCurrentSharePointUserEmail() || localStorage.getItem('mock_user_email') || 'arlenloran@gmail.com';
          if (email.toLowerCase().trim() === curEmail.toLowerCase().trim()) {
            const allowed = await isUserAllowed(curEmail);
            setHasAccess(allowed);
          }
        } else {
          setAccessMessage({ type: 'error', text: 'Erro ao remover acesso' });
        }
      } catch (err: any) {
        setAccessMessage({ type: 'error', text: err?.message || 'Erro ao remover acesso' });
      } finally {
        setIsSavingAccess(false);
      }
    }
  };

  // Form states
  const [sectionTitle, setSectionTitle] = useState('');
  const [metricForm, setMetricForm] = useState({
    title: '',
    objective: '',
    rules: [] as string[],
    sqlQuery: '',
    refreshInterval: 5
  });

  const loadConfig = async () => {
    setIsCheckingAccess(true);
    await ensureSharePointConfig();
    const email = getCurrentSharePointUserEmail() || localStorage.getItem('mock_user_email') || 'arlenloran@gmail.com';
    setUserEmail(email);
    const allowed = await isUserAllowed(email);
    setHasAccess(allowed);
    if (allowed) {
      const data = await fetchDashboardConfig();
      setSections(data);
      const teamsId = await getTeamsChatId();
      setTeamsChatId(teamsId);

      const emailEnabled = await getEmailAlertsEnabled();
      setEmailAlertsEnabled(emailEnabled);

      const teamsEnabled = await getTeamsAlertsEnabled();
      setTeamsAlertsEnabled(teamsEnabled);
    }
    setIsCheckingAccess(false);
    setIsLoading(false);
  };

  const handleToggleEmailAlerts = async (val: boolean) => {
    try {
      const ok = await saveEmailAlertsEnabled(val);
      if (ok) {
        setEmailAlertsEnabled(val);
      } else {
        alert("Erro ao salvar configuração de e-mail no SharePoint");
      }
    } catch (e: any) {
      console.error(e);
      alert(e?.message || "Erro ao salvar configuração");
    }
  };

  const handleToggleTeamsAlerts = async (val: boolean) => {
    if (val && !teamsChatId.trim()) {
      setTeamsMessage({ 
        type: 'error', 
        text: 'Não é possível ativar as notificações do Teams sem antes definir e salvar um ID de chat válido!' 
      });
      setTimeout(() => setTeamsMessage(null), 7000);
      return;
    }
    try {
      const ok = await saveTeamsAlertsEnabled(val);
      if (ok) {
        setTeamsAlertsEnabled(val);
        setTeamsMessage({ 
          type: 'success', 
          text: `Notificações do Teams ${val ? 'ativadas' : 'desativadas'} com sucesso!` 
        });
        setTimeout(() => setTeamsMessage(null), 5000);
      } else {
        setTeamsMessage({ type: 'error', text: 'Erro ao salvar configuração do Teams no SharePoint' });
      }
    } catch (e: any) {
      console.error(e);
      setTeamsMessage({ type: 'error', text: e?.message || "Erro ao salvar configuração" });
    }
  };

  const loadAlertEmails = async () => {
    try {
      const list = await fetchAlertEmails();
      setAlertEmails(list);
    } catch (err) {
      console.error(err);
    }
  };

  useEffect(() => {
    if (isEmailModalOpen) {
      loadAlertEmails();
    }
  }, [isEmailModalOpen]);

  useEffect(() => {
    if (!testTeamsLink.trim()) {
      setExtractedIdResult('');
      return;
    }
    try {
      const decodedLink = decodeURIComponent(testTeamsLink.trim());
      const match = decodedLink.match(/(19:[a-zA-Z0-9_\-@]+(?:\.[a-zA-Z0-9_\-@]+)?)/);
      if (match && match[1]) {
        const cleanId = match[1].split('/')[0].split('?')[0].split('&')[0];
        setExtractedIdResult(cleanId);
      } else {
        setExtractedIdResult('Não foi possível identificar o ID do Teams "19:...". Verifique se o link colado está no formato padrão.');
      }
    } catch (e) {
      setExtractedIdResult('Erro ao decodificar o link fornecido.');
    }
  }, [testTeamsLink]);

  const handleAddAlertEmail = async () => {
    if (!newAlertEmail.trim() || !newAlertEmail.includes('@')) {
      setEmailConfigMessage({ type: 'error', text: 'Insira um e-mail válido' });
      return;
    }
    setIsSavingEmails(true);
    setEmailConfigMessage(null);
    try {
      const ok = await addAlertEmail(newAlertEmail);
      if (ok) {
        setEmailConfigMessage({ type: 'success', text: 'E-mail cadastrado com sucesso!' });
        setNewAlertEmail('');
        await loadAlertEmails();
      } else {
        setEmailConfigMessage({ type: 'error', text: 'Erro ao cadastrar e-mail' });
      }
    } catch (err: any) {
      setEmailConfigMessage({ type: 'error', text: err?.message || 'Erro ao cadastrar e-mail' });
    } finally {
      setIsSavingEmails(false);
    }
  };

  const handleRemoveAlertEmail = async (id: string, email: string) => {
    if (window.confirm(`Deseja realmente remover ${email} da lista de e-mails de alerta?`)) {
      setIsSavingEmails(true);
      setEmailConfigMessage(null);
      try {
        const ok = await removeAlertEmail(id, email);
        if (ok) {
          setEmailConfigMessage({ type: 'success', text: 'E-mail removido com sucesso!' });
          await loadAlertEmails();
        } else {
          setEmailConfigMessage({ type: 'error', text: 'Erro ao remover e-mail' });
        }
      } catch (err: any) {
        setEmailConfigMessage({ type: 'error', text: err?.message || 'Erro ao remover e-mail' });
      } finally {
        setIsSavingEmails(false);
      }
    }
  };

  useEffect(() => {
    loadConfig();
  }, []);

  const handleSaveSection = async () => {
    if (!sectionTitle.trim()) return;
    setIsSaving(true);
    setStatusMessage(null);
    try {
      if (editingSection?.id) {
        await updateDivision(editingSection.id, sectionTitle, 1);
      } else {
        await addDivision(sectionTitle, sections.length + 1);
      }
      setSectionTitle('');
      setEditingSection(null);
      setIsSectionModalOpen(false);
      setStatusMessage({ type: 'success', text: 'Divisão salva com sucesso!' });
      loadConfig();
    } catch (err: any) {
      console.error(err);
      setStatusMessage({ type: 'error', text: `Erro ao salvar divisão: ${err.message}` });
    } finally {
      setIsSaving(false);
    }
  };

  const handleSaveTeamsChatId = async () => {
    setIsSavingTeamsChatId(true);
    setTeamsMessage(null);
    try {
      const ok = await saveTeamsChatId(teamsChatId);
      if (ok) {
        setTeamsMessage({ type: 'success', text: 'ID do grupo Teams atualizado com sucesso!' });
        setTimeout(() => setTeamsMessage(null), 5000);
      } else {
        setTeamsMessage({ type: 'error', text: 'Ocorreu um erro ao salvar o ID no SharePoint' });
      }
    } catch (err: any) {
      setTeamsMessage({ type: 'error', text: err?.message || 'Erro ao conectar ao SharePoint.' });
    } finally {
      setIsSavingTeamsChatId(false);
    }
  };

  const handleSaveMetric = async () => {
    if (!metricForm.title.trim() || !editingMetric?.divisionId) return;
    setIsSaving(true);
    setStatusMessage(null);
    try {
      if (editingMetric.metric.id !== 'new') {
        await updateMetric(editingMetric.metric.id, metricForm);
      } else {
        await addMetric(editingMetric.divisionId, metricForm);
      }
      setEditingMetric(null);
      setIsMetricModalOpen(false);
      setStatusMessage({ type: 'success', text: 'Card salvo com sucesso!' });
      loadConfig();
    } catch (err: any) {
      console.error(err);
      setStatusMessage({ type: 'error', text: `Erro ao salvar card: ${err.message}` });
    } finally {
      setIsSaving(false);
    }
  };

  const openNewMetricModal = (divisionId: string) => {
    setMetricForm({
      title: '',
      objective: '',
      rules: [],
      sqlQuery: '',
      refreshInterval: 5
    });
    setEditingMetric({ 
      metric: { id: 'new', title: '', value: 0, status: 'ok', lastUpdate: '', isDynamic: true, details: [], history: [], rules: [] }, 
      divisionId: divisionId
    });
    setIsMetricModalOpen(true);
  };

  if (isCheckingAccess) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center py-20 gap-4">
        <Activity className="w-12 h-12 text-slate-400 animate-spin" />
        <p className="font-black text-slate-400 uppercase tracking-widest text-[10px]">Verificando Permissões...</p>
      </div>
    );
  }

  if (!hasAccess) {
    return (
      <div className="min-h-screen bg-slate-50 text-slate-900 p-4 md:p-8 flex flex-col items-center justify-center">
        <motion.div 
          initial={{ opacity: 0, y: 30 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white border border-slate-200 p-8 rounded-3xl shadow-xl w-full max-w-md text-center"
        >
          <div className="mx-auto w-16 h-16 bg-red-50 text-brand-red rounded-2xl flex items-center justify-center mb-6">
            <Lock className="w-8 h-8 animate-bounce" />
          </div>
          
          <h1 className="text-2xl font-black uppercase italic tracking-tighter text-slate-900">Acesso Restrito</h1>
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mt-1 mb-6">Painel Administrativo bloqueado</p>
          
          <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 text-left mb-6">
            <p className="text-xs font-bold text-slate-500 uppercase tracking-tight mb-1">E-mail Identificado:</p>
            <p className="text-xs font-black text-slate-800 break-all">{userEmail || 'Nenhum e-mail identificado'}</p>
          </div>

          <p className="text-xs text-slate-500 leading-relaxed mb-8">
            Desculpe, apenas usuários cadastrados na lista de permissões possuem autorização para gerenciar a estrutura das métricas.
          </p>

          <div className="flex flex-col gap-3">
            <a href="/" className="w-full py-4 bg-slate-900 text-white rounded-2xl font-black uppercase tracking-widest text-xs flex items-center justify-center gap-2 hover:bg-slate-800 transition-all active:scale-95 shadow-md">
              <ArrowLeft className="w-4 h-4" /> Voltar ao Dashboard
            </a>
            
            {!hasSpContext() && (
              <button 
                onClick={() => {
                  const val = window.prompt("Simular novo e-mail para teste:", userEmail);
                  if (val) {
                    localStorage.setItem('mock_user_email', val.trim());
                    window.location.reload();
                  }
                }}
                className="w-full py-3 bg-white border border-slate-200 text-slate-600 rounded-2xl font-black uppercase tracking-widest text-[9px] hover:bg-slate-50 transition-all cursor-pointer"
              >
                Simular outro e-mail (modo teste)
              </button>
            )}
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 p-4 sm:p-8">
      <header className="max-w-6xl mx-auto flex flex-col md:flex-row justify-between items-start md:items-center gap-6 mb-10 md:mb-12">
        <div className="flex items-center gap-3 sm:gap-4">
          <div className="p-2.5 sm:p-3 bg-slate-900 rounded-2xl text-white">
            <Settings className="w-6 h-6 sm:w-8 sm:h-8" />
          </div>
          <div>
            <h1 className="text-2xl sm:text-3xl font-black uppercase italic tracking-tighter">Painel administrativo</h1>
            <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mt-1">Gerenciamento de Estrutura & Métricas</p>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2.5 sm:gap-3 w-full md:w-auto">
          <button 
            onClick={() => setIsAccessModalOpen(true)} 
            className="flex-grow sm:flex-grow-0 flex items-center justify-center gap-2 px-4 py-2.5 bg-slate-900 text-white rounded-xl font-black text-[10px] uppercase tracking-widest hover:bg-slate-800 transition-all cursor-pointer"
          >
            <Shield className="w-4 h-4 text-brand-red animate-pulse" /> Gerenciar Acessos
          </button>
          <a href="/" className="flex-grow sm:flex-grow-0 flex items-center justify-center gap-2 px-4 py-2.5 bg-white border border-slate-200 rounded-xl font-black text-[10px] uppercase tracking-widest hover:bg-slate-50 transition-all">
            <ArrowLeft className="w-4 h-4" /> Voltar ao Dashboard
          </a>
        </div>
      </header>

      <main className="max-w-6xl mx-auto space-y-6 sm:space-y-8 pb-32">
        {/* Painel de Integração Teams */}
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4 sm:p-6">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              <div className="p-2 sm:p-3 bg-[#e6f0fa] text-[#1f4e79] rounded-xl flex-shrink-0">
                <Shield className="w-5 h-5 sm:w-6 sm:h-6" />
              </div>
              <div className="min-w-0">
                <h3 className="text-sm sm:text-base font-black uppercase text-slate-900 tracking-tight flex items-center gap-2 flex-wrap">
                  Notificações do Microsoft Teams 
                  <span className="text-[9px] font-black uppercase bg-[#1f4e79] text-white px-2 py-0.5 rounded-full">Integração</span>
                  <button 
                    type="button"
                    onClick={() => setIsHelpModalOpen(true)}
                    className="inline-flex items-center gap-1 px-2 py-0.5 bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 rounded-lg text-[9px] font-black uppercase tracking-wider transition-all cursor-pointer shadow-xs ml-1"
                    title="Aprenda a capturar o ID correto"
                  >
                    <HelpCircle className="w-3 h-3 text-blue-600" />
                    <span>Como obter o ID?</span>
                  </button>
                </h3>
                <p className="text-[11px] sm:text-xs text-slate-500 mt-1">Defina o ID do grupo/chat do Teams de destino (lista SharePoint <strong>App_Dash_Configs</strong>).</p>
              </div>
            </div>
            
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-4 w-full md:w-auto">
              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2 flex-grow">
                <div className="flex-grow">
                  <input 
                    type="text" 
                    value={teamsChatId}
                    onChange={(e) => setTeamsChatId(e.target.value)}
                    placeholder="ID do Chat do Teams"
                    className="w-full px-4 py-2.5 bg-slate-50 border border-slate-200 text-slate-950 font-mono text-xs font-bold rounded-xl outline-none focus:ring-2 focus:ring-slate-950 transition-all font-mono"
                  />
                </div>
                <button 
                  onClick={handleSaveTeamsChatId}
                  disabled={isSavingTeamsChatId}
                  className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white rounded-xl font-black text-xs uppercase tracking-widest transition-all cursor-pointer whitespace-nowrap flex items-center justify-center gap-1.5 disabled:opacity-50"
                >
                  <Save className="w-4 h-4" /> {isSavingTeamsChatId ? 'Salvando...' : 'Salvar ID'}
                </button>
              </div>

              {/* Status Toggle for Teams */}
              <div className="flex items-center justify-between sm:justify-end gap-3 border-t sm:border-t-0 border-slate-100 pt-3 sm:pt-0 sm:pl-4 sm:border-l">
                <div className="flex flex-col">
                  <span className="text-xs font-black uppercase text-slate-600">Ativar Teams</span>
                  <span className="text-[10px] text-slate-400">Canal ativo de alertas</span>
                </div>
                <button
                  onClick={() => handleToggleTeamsAlerts(!teamsAlertsEnabled)}
                  className={`relative inline-flex h-6 w-12 items-center rounded-full transition-colors duration-300 ${
                    teamsAlertsEnabled ? 'bg-[#1f4e79]' : 'bg-slate-200'
                  }`}
                >
                  <span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform duration-300 ${
                    teamsAlertsEnabled ? 'translate-x-[26px]' : 'translate-x-[4px]'
                  }`} />
                </button>
              </div>
            </div>
          </div>

          <AnimatePresence>
            {teamsMessage && (
              <motion.div 
                initial={{ opacity: 0, height: 0 }}
                animate={{ opacity: 1, height: 'auto' }}
                exit={{ opacity: 0, height: 0 }}
                className={`mt-4 p-3 rounded-xl border font-bold text-xs uppercase tracking-wide flex items-center gap-2 ${
                  teamsMessage.type === 'success' 
                    ? 'bg-emerald-50 border-emerald-100 text-emerald-600' 
                    : 'bg-red-50 border-red-100 text-red-600'
                }`}
              >
                <Shield className="w-4 h-4" />
                {teamsMessage.text}
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Painel de Configuração de E-mail */}
        <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4 sm:p-6">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              <div className="p-2 sm:p-3 bg-red-50 text-brand-red rounded-xl flex-shrink-0">
                <Mail className="w-5 h-5 sm:w-6 sm:h-6" />
              </div>
              <div className="min-w-0">
                <h3 className="text-sm sm:text-base font-black uppercase text-slate-900 tracking-tight flex items-center gap-2">
                  Notificações por E-mail
                  <span className="text-[9px] font-black uppercase bg-brand-red text-white px-2 py-0.5 rounded-full">Integração</span>
                </h3>
                <p className="text-[11px] sm:text-xs text-slate-500 mt-1">
                  Ative o envio de divergências e configure quem receberá os alertas (lista SharePoint <strong>App_Dash_Emails</strong>).
                </p>
              </div>
            </div>
            
            <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-4 w-full md:w-auto">
              {/* Register Button */}
              <button
                onClick={() => { setIsEmailModalOpen(true); }}
                className="px-4 py-2.5 bg-slate-100 hover:bg-slate-200 border border-slate-200 rounded-xl font-black text-xs uppercase tracking-widest text-slate-700 transition-all cursor-pointer whitespace-nowrap flex items-center justify-center gap-1.5"
              >
                <Plus className="w-4 h-4 text-slate-500" /> Cadastrar E-mails
              </button>

              {/* Status Toggle for Email */}
              <div className="flex items-center justify-between sm:justify-end gap-3 border-t sm:border-t-0 border-slate-100 pt-3 sm:pt-0 sm:pl-4 sm:border-l overflow-visible">
                <div className="flex flex-col">
                  <span className="text-xs font-black uppercase text-slate-600 font-bold">Ativar E-mail</span>
                  <span className="text-[10px] text-slate-400">Canal ativo de alertas</span>
                </div>
                <button
                  onClick={() => handleToggleEmailAlerts(!emailAlertsEnabled)}
                  className={`relative inline-flex h-6 w-12 items-center rounded-full transition-colors duration-300 ${
                    emailAlertsEnabled ? 'bg-[#1f4e79]' : 'bg-slate-200'
                  }`}
                >
                  <span className={`inline-block h-4 w-4 transform rounded-full bg-white transition-transform duration-300 ${
                    emailAlertsEnabled ? 'translate-x-[26px]' : 'translate-x-[4px]'
                  }`} />
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
          <h2 className="text-lg sm:text-xl font-black uppercase italic tracking-tight flex items-center gap-2">
            <Layout className="w-5 h-5 text-brand-red" /> Divisões do Dashboard
          </h2>
          <button 
            onClick={() => { setEditingSection(null); setSectionTitle(''); setIsSectionModalOpen(true); }}
            className="w-full sm:w-auto px-6 py-3 bg-slate-900 text-white rounded-xl font-black text-xs uppercase tracking-widest flex items-center justify-center gap-2 hover:bg-slate-800 transition-all shadow-lg shadow-slate-200 active:scale-95"
          >
            <Plus className="w-4 h-4" /> Nova Divisão
          </button>
        </div>

        <AnimatePresence>
          {statusMessage && (
            <motion.div 
              initial={{ opacity: 0, y: -20 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -20 }}
              className={`p-4 rounded-xl border flex items-center gap-3 font-bold text-xs uppercase tracking-widest ${
                statusMessage.type === 'success' 
                  ? 'bg-emerald-50 border-emerald-100 text-emerald-600' 
                  : 'bg-red-50 border-red-100 text-red-600'
              }`}
            >
              {statusMessage.type === 'success' ? <Shield className="w-5 h-5" /> : <AlertTriangle className="w-5 h-5" />}
              {statusMessage.text}
              <button onClick={() => setStatusMessage(null)} className="ml-auto opacity-50 hover:opacity-100"><X className="w-4 h-4" /></button>
            </motion.div>
          )}
        </AnimatePresence>

        {isLoading ? (
          <div className="flex flex-col items-center justify-center py-20 gap-4">
            <Activity className="w-12 h-12 text-slate-300 animate-spin" />
            <p className="font-black text-slate-400 uppercase tracking-widest text-[10px]">Carregando Configurações...</p>
          </div>
        ) : (
          <div className="grid gap-6">
            {sections.map((section, sIdx) => (
              <motion.div 
                layout
                key={section.id || section.title}
                draggable
                onDragStart={(e) => {
                  // Only drag section if we are dragging the handle or general header area
                  setDraggedSectionIndex(sIdx);
                }}
                onDragOver={(e) => {
                  e.preventDefault();
                }}
                onDrop={(e) => {
                  if (draggedSectionIndex !== null && draggedSectionIndex !== sIdx) {
                    handleMoveSection(draggedSectionIndex, sIdx);
                  }
                  setDraggedSectionIndex(null);
                }}
                className={`bg-white border rounded-2xl overflow-hidden shadow-sm hover:shadow-md transition-all duration-300 ${
                  draggedSectionIndex === sIdx ? 'opacity-40 border-dashed border-brand-red scale-[0.99]' : 'border-slate-200'
                }`}
              >
                <div className="p-4 sm:p-6 border-b border-slate-100 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 bg-slate-50/50">
                  <div className="flex items-center gap-3 sm:gap-4 w-full sm:w-auto">
                    <div 
                      className="text-slate-400 cursor-grab active:cursor-grabbing hover:text-slate-600 transition-colors p-1 flex-shrink-0" 
                      title="Arraste para reordenar divisão"
                    >
                      <GripVertical className="w-5 h-5" />
                    </div>
                    <div className="w-1.5 h-6 sm:w-2 sm:h-8 bg-brand-red rounded-full flex-shrink-0" />
                    <h3 className="text-base sm:text-lg font-black uppercase italic tracking-tight truncate">{section.title}</h3>
                  </div>
                  <div className="flex gap-2 items-center justify-between sm:justify-end w-full sm:w-auto border-t sm:border-t-0 pt-3 sm:pt-0 border-slate-100">
                    <div className="flex border border-slate-200 rounded-lg overflow-hidden bg-white mr-1 flex-shrink-0">
                      <button 
                        onClick={() => handleMoveSection(sIdx, sIdx - 1)}
                        disabled={sIdx === 0}
                        className="p-1 px-2.5 bg-white text-slate-400 hover:text-slate-800 disabled:opacity-30 hover:bg-slate-50 transition-all border-r border-slate-150 cursor-pointer"
                        title="Mover divisão para cima"
                      >
                        <ChevronUp className="w-4 h-4" />
                      </button>
                      <button 
                        onClick={() => handleMoveSection(sIdx, sIdx + 1)}
                        disabled={sIdx === sections.length - 1}
                        className="p-1 px-2.5 bg-white text-slate-400 hover:text-slate-800 disabled:opacity-30 hover:bg-slate-50 transition-all cursor-pointer"
                        title="Mover divisão para baixo"
                      >
                        <ChevronDown className="w-4 h-4" />
                      </button>
                    </div>
                    <div className="flex gap-1">
                      <button 
                        onClick={() => { setEditingSection(section); setSectionTitle(section.title); setIsSectionModalOpen(true); }}
                        className="p-2 text-slate-400 hover:text-slate-900 hover:bg-white rounded-lg transition-all"
                        title="Editar divisão"
                      >
                        <Edit2 className="w-5 h-5" />
                      </button>
                      <button 
                        onClick={async () => {
                          if (section.id && window.confirm(`Deseja realmente excluir a divisão "${section.title}"?`)) {
                            try {
                              setIsSaving(true);
                              await deleteDivision(section.id);
                              setStatusMessage({ type: 'success', text: 'Divisão excluída!' });
                              loadConfig();
                            } catch (err: any) {
                              setStatusMessage({ type: 'error', text: `Erro ao excluir: ${err.message}` });
                            } finally {
                              setIsSaving(false);
                            }
                          }
                        }}
                        className="p-2 text-slate-400 hover:text-red-500 hover:bg-white rounded-lg transition-all"
                        title="Excluir divisão"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    </div>
                  </div>
                </div>

                <div className="p-4 sm:p-6">
                  <div className="grid gap-4">
                    {section.metrics.map((metric, mIdx) => (
                      <div 
                        key={metric.id} 
                        draggable
                        onDragStart={(e) => {
                          e.stopPropagation();
                          setDraggedMetricIndex({ sectionId: section.id!, index: mIdx });
                        }}
                        onDragOver={(e) => {
                          e.preventDefault();
                        }}
                        onDrop={(e) => {
                          e.stopPropagation();
                          if (draggedMetricIndex !== null && draggedMetricIndex.sectionId === section.id! && draggedMetricIndex.index !== mIdx) {
                            handleMoveMetric(section.id!, draggedMetricIndex.index, mIdx);
                          }
                          setDraggedMetricIndex(null);
                        }}
                        className={`flex flex-col sm:flex-row justify-between items-start sm:items-center p-3.5 sm:p-4 bg-slate-50 rounded-xl border gap-4 group transition-all duration-300 ${
                          draggedMetricIndex && draggedMetricIndex.sectionId === section.id! && draggedMetricIndex.index === mIdx
                            ? 'opacity-40 border-dashed border-brand-red scale-[0.99] border-brand-red bg-red-50/5'
                            : 'border-slate-100'
                        }`}
                      >
                        <div className="flex items-center gap-2.5 sm:gap-4 w-full sm:w-auto">
                          <div 
                            className="text-slate-400 cursor-grab active:cursor-grabbing hover:text-slate-600 transition-colors p-1" 
                            title="Arraste para reordenar card"
                          >
                            <GripVertical className="w-4 h-4" />
                          </div>
                          <div className="p-2 sm:p-3 bg-white border border-slate-200 rounded-xl group-hover:border-brand-red transition-colors flex-shrink-0">
                            <Activity className="w-4 h-4 sm:w-5 sm:h-5 text-slate-400 group-hover:text-brand-red" />
                          </div>
                          <div className="min-w-0 flex-grow">
                            <h4 className="font-black text-slate-800 uppercase text-xs truncate">{metric.title}</h4>
                            <div className="flex gap-3 sm:gap-4 mt-1">
                              <span className="text-[9px] font-bold text-slate-400 uppercase flex items-center gap-1"><Clock className="w-2.5 h-2.5" /> a cada {metric.refreshInterval}m</span>
                              <span className="text-[9px] font-bold text-slate-400 uppercase flex items-center gap-1"><Shield className="w-2.5 h-2.5" /> {metric.rules?.length || 0} regras</span>
                            </div>
                          </div>
                        </div>
                        <div className="flex flex-wrap gap-2 items-center justify-between sm:justify-end w-full sm:w-auto border-t sm:border-t-0 pt-3 sm:pt-0 border-slate-200/60">
                          <div className="flex border border-slate-200 rounded-lg overflow-hidden bg-white">
                            <button 
                              onClick={() => handleMoveMetric(section.id!, mIdx, mIdx - 1)}
                              disabled={mIdx === 0}
                              className="p-1 px-2.5 bg-white text-slate-400 hover:text-slate-800 disabled:opacity-30 hover:bg-slate-50 transition-all border-r border-slate-150 cursor-pointer"
                              title="Mover card para cima"
                            >
                              <ChevronUp className="w-3.5 h-3.5" />
                            </button>
                            <button 
                              onClick={() => handleMoveMetric(section.id!, mIdx, mIdx + 1)}
                              disabled={mIdx === section.metrics.length - 1}
                              className="p-1 px-2.5 bg-white text-slate-400 hover:text-slate-800 disabled:opacity-30 hover:bg-slate-50 transition-all cursor-pointer"
                              title="Mover card para baixo"
                            >
                              <ChevronDown className="w-3.5 h-3.5" />
                            </button>
                          </div>
                          <div className="flex gap-1.5 ml-auto sm:ml-0">
                            <button 
                              onClick={() => {
                                setMetricForm({
                                  title: metric.title,
                                  objective: metric.objective || '',
                                  rules: metric.rules || [],
                                  sqlQuery: metric.sqlQuery || '',
                                  refreshInterval: metric.refreshInterval || 5
                                });
                                setEditingMetric({ metric, divisionId: section.id || '' });
                                setIsMetricModalOpen(true);
                              }}
                              className="px-2.5 sm:px-4 py-2 bg-white border border-slate-200 text-slate-700 rounded-lg font-black text-[10px] uppercase tracking-widest hover:bg-slate-50 transition-all flex items-center gap-1.5"
                            >
                              <Edit2 className="w-3 h-3" /> Editar
                            </button>
                            <button 
                              onClick={async () => {
                                if (window.confirm(`Deseja realmente excluir o card "${metric.title}"?`)) {
                                  try {
                                    setIsSaving(true);
                                    await deleteMetric(metric.id);
                                    setStatusMessage({ type: 'success', text: 'Card excluído!' });
                                    loadConfig();
                                  } catch (err: any) {
                                    setStatusMessage({ type: 'error', text: `Erro ao excluir: ${err.message}` });
                                  } finally {
                                    setIsSaving(false);
                                  }
                                }
                              }}
                              className="px-2.5 sm:px-4 py-2 bg-white border border-slate-200 text-red-500 rounded-lg font-black text-[10px] uppercase tracking-widest hover:border-red-100 hover:bg-red-50 transition-all flex items-center gap-1.5"
                            >
                              <Trash2 className="w-3 h-3" /> Excluir
                            </button>
                          </div>
                        </div>
                      </div>
                    ))}
                    <button 
                      onClick={() => section.id && openNewMetricModal(section.id)}
                      className="w-full py-4 border-2 border-dashed border-slate-200 rounded-xl flex items-center justify-center gap-2 text-slate-400 hover:border-slate-400 hover:text-slate-600 transition-all group cursor-pointer"
                    >
                      <Plus className="w-5 h-5 group-hover:scale-125 transition-transform" />
                      <span className="font-black uppercase text-xs tracking-widest">Adicionar Novo Card</span>
                    </button>
                  </div>
                </div>
              </motion.div>
            ))}
          </div>
        )}
      </main>

      {/* Division Modal */}
      <AnimatePresence>
        {isSectionModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.9 }}
              className="bg-white rounded-2xl sm:rounded-3xl p-5 sm:p-8 w-full max-w-md shadow-2xl"
            >
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl sm:text-2xl font-black uppercase italic tracking-tighter">{editingSection ? 'Editar Divisão' : 'Nova Divisão'}</h3>
                <button onClick={() => setIsSectionModalOpen(false)} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><X className="w-5 h-5 sm:w-6 sm:h-6" /></button>
              </div>
              <div className="space-y-6">
                <div>
                  <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2">Título da Divisão</label>
                  <input 
                    type="text" 
                    value={sectionTitle}
                    onChange={(e) => setSectionTitle(e.target.value)}
                    placeholder="Ex: Qualidade Operacional"
                    className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-slate-900 transition-all font-bold text-sm"
                  />
                </div>
                <button 
                  onClick={handleSaveSection}
                  disabled={isSaving}
                  className="w-full py-3.5 sm:py-4 bg-slate-900 text-white rounded-xl sm:rounded-2xl font-black uppercase tracking-widest text-xs sm:text-sm flex items-center justify-center gap-2 hover:bg-slate-800 transition-all disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Save className={isSaving ? "w-4 h-4 sm:w-5 sm:h-5 animate-spin" : "w-4 h-4 sm:w-5 sm:h-5"} /> {isSaving ? 'Salvando...' : (editingSection ? 'Salvar Alterações' : 'Criar Divisão')}
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Metric Modal */}
      <AnimatePresence>
        {isMetricModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, y: 30 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: 30 }}
              className="bg-white rounded-2xl sm:rounded-3xl w-full max-w-2xl max-h-[92vh] overflow-hidden flex flex-col shadow-2xl"
            >
              <header className="p-5 sm:p-8 border-b border-slate-100 flex justify-between items-center bg-slate-50/50">
                <div>
                  <h3 className="text-xl sm:text-2xl font-black uppercase italic tracking-tighter">
                    {editingMetric?.metric.id === 'new' ? 'Novo Card de Métrica' : 'Editar Card de Métrica'}
                  </h3>
                  <p className="text-[10px] font-black text-slate-400 uppercase mt-0.5">Configuração dinâmica e persistência</p>
                </div>
                <button onClick={() => setIsMetricModalOpen(false)} className="p-2 hover:bg-slate-200 rounded-full transition-colors"><X className="w-5 h-5 sm:w-6 sm:h-6" /></button>
              </header>

              <div className="flex-grow overflow-y-auto p-5 sm:p-8 space-y-6 sm:space-y-8">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-5 sm:gap-6">
                  <div className="space-y-5 sm:space-y-6">
                    <div>
                      <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 flex items-center gap-2"><Activity className="w-3 h-3" /> Título do Card</label>
                      <input 
                        type="text" 
                        value={metricForm.title}
                        onChange={(e) => setMetricForm({...metricForm, title: e.target.value})}
                        className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-slate-900 transition-all font-bold text-sm"
                        placeholder="Ex: Divergência de Inventário"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 flex items-center gap-2"><Clock className="w-3 h-3" /> Intervalo de Atualização (Minutos)</label>
                      <input 
                        type="number" 
                        value={metricForm.refreshInterval}
                        onChange={(e) => setMetricForm({...metricForm, refreshInterval: parseInt(e.target.value)})}
                        className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-slate-900 transition-all font-bold text-sm"
                      />
                    </div>
                  </div>

                  <div>
                    <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 flex items-center gap-2"><BookOpen className="w-3 h-3" /> Objetivo (Descrição)</label>
                    <textarea 
                      value={metricForm.objective}
                      onChange={(e) => setMetricForm({...metricForm, objective: e.target.value})}
                      className="w-full h-[100px] md:h-[135px] px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-slate-900 transition-all font-bold resize-none text-sm"
                      placeholder="Descreva o propósito desta métrica..."
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 flex items-center gap-2"><Database className="w-3 h-3" /> Query SQL (Fonte de Dados)</label>
                  <textarea 
                    value={metricForm.sqlQuery}
                    onChange={(e) => setMetricForm({...metricForm, sqlQuery: e.target.value})}
                    className="w-full h-[120px] sm:h-[150px] px-4 py-3 bg-slate-950 text-emerald-500 font-mono text-xs sm:text-sm border-2 border-slate-800 rounded-xl outline-none focus:border-brand-red transition-all resize-none shadow-inner"
                    placeholder="SELECT * FROM TABELA WHERE..."
                  />
                  <div className="mt-2 text-[9px] text-slate-400 font-bold uppercase p-3 bg-slate-100 rounded-lg flex items-start gap-2">
                    <AlertTriangle className="w-3.5 h-3.5 text-brand-red flex-shrink-0" /> 
                    <span>Atuando em tempo real sobre a base do ERP selecionada nas configurações.</span>
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-black uppercase tracking-widest text-slate-400 mb-2 flex items-center gap-2"><Shield className="w-3 h-3" /> Regras de Negócio</label>
                  <div className="space-y-2">
                    {metricForm.rules.map((rule, idx) => (
                      <div key={idx} className="flex gap-2">
                        <input 
                          type="text" 
                          value={rule}
                          onChange={(e) => {
                            const newRules = [...metricForm.rules];
                            newRules[idx] = e.target.value;
                            setMetricForm({...metricForm, rules: newRules});
                          }}
                          className="flex-grow px-3 sm:px-4 py-2.5 sm:py-3 bg-slate-50 border border-slate-200 rounded-xl outline-none focus:ring-2 focus:ring-slate-900 transition-all font-bold text-xs sm:text-sm"
                        />
                        <button onClick={() => setMetricForm({...metricForm, rules: metricForm.rules.filter((_, i) => i !== idx)})} className="p-2.5 text-red-500 hover:bg-red-50 rounded-xl transition-all cursor-pointer"><Trash2 className="w-4 h-4 sm:w-5 sm:h-5" /></button>
                      </div>
                    ))}
                    <button 
                      onClick={() => setMetricForm({...metricForm, rules: [...metricForm.rules, '']})}
                      className="w-full py-2.5 sm:py-3 border-2 border-dashed border-slate-200 rounded-xl text-slate-400 hover:border-slate-400 hover:text-slate-600 transition-all font-black uppercase text-[10px] tracking-widest cursor-pointer"
                    >
                      + Adicionar Regra
                    </button>
                  </div>
                </div>
              </div>

              <footer className="p-5 sm:p-8 border-t border-slate-100 bg-slate-50/50 flex flex-col-reverse sm:flex-row justify-end gap-2.5 sm:gap-3">
                <button onClick={() => setIsMetricModalOpen(false)} disabled={isSaving} className="w-full sm:w-auto px-6 sm:px-8 py-3 bg-white border border-slate-200 text-slate-600 rounded-xl sm:rounded-2xl font-black uppercase tracking-widest text-xs hover:bg-slate-100 transition-all disabled:opacity-50">Cancelar</button>
                <button 
                  onClick={handleSaveMetric}
                  disabled={isSaving}
                  className="w-full sm:w-auto px-8 sm:px-10 py-3 bg-slate-900 text-white rounded-xl sm:rounded-2xl font-black uppercase tracking-widest text-xs flex items-center justify-center gap-2 hover:bg-slate-800 transition-all shadow-lg active:scale-95 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  <Save className={isSaving ? "w-4 h-4 animate-spin" : "w-4 h-4"} /> {isSaving ? 'Salvando...' : 'Salvar Card'}
                </button>
              </footer>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Modal de Gerenciamento de Acessos */}
      <AnimatePresence>
        {isAccessModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="rounded-2xl sm:rounded-3xl p-5 sm:p-8 w-full max-w-lg shadow-2xl bg-white border border-slate-200 text-slate-900 overflow-hidden flex flex-col"
            >
              <div className="flex justify-between items-start gap-4 mb-6">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="p-2 rounded-xl bg-red-50 text-brand-red flex-shrink-0">
                    <Shield className="w-5 h-5 sm:w-6 sm:h-6" />
                  </div>
                  <div>
                    <h3 className="text-lg sm:text-xl font-black uppercase italic tracking-tighter leading-tight">Controle de acessos</h3>
                    <p className="text-[9px] font-black uppercase tracking-wider text-slate-400 mt-0.5">Gerenciar e-mails permitidos</p>
                  </div>
                </div>
                <button 
                  onClick={() => setIsAccessModalOpen(false)} 
                  className="p-1.5 rounded-full transition-colors hover:bg-slate-100 text-slate-400 hover:text-slate-900 flex-shrink-0"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {accessMessage && (
                <div className={`p-3 rounded-xl text-[10px] sm:text-xs font-bold uppercase tracking-wider mb-4 border ${accessMessage.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-600' : 'bg-red-50 border-red-100 text-red-600'}`}>
                  {accessMessage.text}
                </div>
              )}

              {/* Form to Add User */}
              <div className="mb-6">
                <label className="block text-[9px] font-black uppercase tracking-widest mb-1.5 text-slate-400">Conceder Novo Acesso (E-mail)</label>
                <div className="flex flex-col sm:flex-row gap-2">
                  <input 
                    type="email" 
                    value={newAccessEmail}
                    onChange={(e) => setNewAccessEmail(e.target.value)}
                    placeholder="Ex: usuario@empresa.com"
                    className="flex-grow px-4 py-2.5 sm:py-3 border rounded-xl outline-none transition-all font-bold text-xs bg-slate-50 border-slate-200 text-slate-900 focus:ring-1 focus:ring-slate-900"
                  />
                  <button 
                    onClick={handleAddAccess}
                    disabled={isSavingAccess}
                    className="px-5 py-2.5 sm:py-3 bg-brand-red hover:bg-red-650 text-white rounded-xl font-black text-xs uppercase tracking-widest transition-all active:scale-95 flex items-center justify-center gap-1.5"
                  >
                    Adicionar
                  </button>
                </div>
              </div>

              {/* Scrollable list of active users */}
              <div className="flex flex-col">
                <label className="block text-[9px] font-black uppercase tracking-widest mb-2 text-slate-400">Usuários com Permissão ({allowedUsers.length})</label>
                <div className="max-h-[180px] sm:max-h-[220px] overflow-y-auto rounded-xl sm:rounded-2xl border bg-slate-50 border-slate-200">
                  {allowedUsers.length === 0 ? (
                    <p className="text-[10px] uppercase font-bold text-center py-8 text-slate-400 tracking-wider">Nenhum e-mail cadastrado</p>
                  ) : (
                    <div className="divide-y divide-slate-100">
                      {allowedUsers.map(u => (
                        <div key={u.id} className="p-3 sm:p-3.5 flex justify-between items-center hover:bg-slate-100 transition-colors">
                          <div className="flex items-center gap-2 max-w-[80%] min-w-0">
                            <div className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-emerald-500 flex-shrink-0" />
                            <span className="font-extrabold text-[11px] sm:text-xs tracking-tight select-all truncate">{u.email}</span>
                          </div>
                          {(getCurrentSharePointUserEmail() || 'arlenloran@gmail.com').toLowerCase().trim() !== u.email.toLowerCase().trim() && (
                            <button 
                              onClick={() => handleRemoveAccess(u.id, u.email)}
                              disabled={isSavingAccess}
                              className="p-1.5 text-slate-400 hover:text-brand-red rounded-lg hover:bg-red-50/10 transition-colors flex-shrink-0"
                              title="Remover Permissão"
                            >
                              <X className="w-4 h-4" />
                            </button>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Modal de Gerenciamento de E-mails de Alerta */}
      <AnimatePresence>
        {isEmailModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="rounded-2xl sm:rounded-3xl p-5 sm:p-8 w-full max-w-lg shadow-2xl bg-white border border-slate-200 text-slate-900 overflow-hidden flex flex-col"
            >
              <div className="flex justify-between items-start gap-4 mb-6">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="p-2 rounded-xl bg-red-50 text-brand-red flex-shrink-0">
                    <Mail className="w-5 h-5 sm:w-6 sm:h-6" />
                  </div>
                  <div>
                    <h3 className="text-lg sm:text-xl font-black uppercase italic tracking-tighter leading-tight">Lista de Alerta</h3>
                    <p className="text-[9px] font-black uppercase tracking-wider text-slate-400 mt-0.5">Gerenciar destinatários dos e-mails</p>
                  </div>
                </div>
                <button 
                  onClick={() => { setIsEmailModalOpen(false); setEmailConfigMessage(null); }} 
                  className="p-1.5 rounded-full transition-colors hover:bg-slate-100 text-slate-400 hover:text-slate-900 flex-shrink-0"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {emailConfigMessage && (
                <div className={`p-3 rounded-xl text-[10px] sm:text-xs font-bold uppercase tracking-wider mb-4 border ${emailConfigMessage.type === 'success' ? 'bg-emerald-50 border-emerald-100 text-emerald-600' : 'bg-red-50 border-red-100 text-red-600'}`}>
                  {emailConfigMessage.text}
                </div>
              )}

              {/* Form to Add User Email */}
              <div className="mb-6">
                <label className="block text-[9px] font-black uppercase tracking-widest mb-1.5 text-slate-400">Cadastrar E-mail de Alerta</label>
                <div className="flex flex-col sm:flex-row gap-2">
                  <input 
                    type="email" 
                    value={newAlertEmail}
                    onChange={(e) => setNewAlertEmail(e.target.value)}
                    placeholder="Ex: administrador@empresa.com"
                    className="flex-grow px-4 py-2.5 sm:py-3 border rounded-xl outline-none transition-all font-bold text-xs bg-slate-50 border-slate-200 text-slate-900 focus:ring-1 focus:ring-slate-900"
                  />
                  <button 
                    onClick={handleAddAlertEmail}
                    disabled={isSavingEmails}
                    className="px-5 py-2.5 sm:py-3 bg-brand-red hover:bg-red-650 text-white rounded-xl font-black text-xs uppercase tracking-widest transition-all active:scale-95 flex items-center justify-center gap-1.5 whitespace-nowrap disabled:opacity-50"
                  >
                    Adicionar
                  </button>
                </div>
              </div>

              {/* Scrollable list of emails */}
              <div className="flex flex-col">
                <label className="block text-[9px] font-black uppercase tracking-widest mb-2 text-slate-400">Destinatários Cadastrados ({alertEmails.length})</label>
                <div className="max-h-[180px] sm:max-h-[220px] overflow-y-auto rounded-xl sm:rounded-2xl border bg-slate-50 border-slate-200">
                  {alertEmails.length === 0 ? (
                    <p className="text-[10px] uppercase font-bold text-center py-8 text-slate-400 tracking-wider">Nenhum e-mail cadastrado</p>
                  ) : (
                    <div className="divide-y divide-slate-100">
                      {alertEmails.map(u => (
                        <div key={u.id} className="p-3 sm:p-3.5 flex justify-between items-center hover:bg-slate-100 transition-colors">
                          <div className="flex items-center gap-2 max-w-[80%] min-w-0">
                            <div className="w-1.5 h-1.5 sm:w-2 sm:h-2 rounded-full bg-blue-500 flex-shrink-0" />
                            <span className="font-extrabold text-[11px] sm:text-xs tracking-tight select-all truncate">{u.email}</span>
                          </div>
                          <button 
                            onClick={() => handleRemoveAlertEmail(u.id, u.email)}
                            disabled={isSavingEmails}
                            className="p-1.5 text-slate-400 hover:text-brand-red rounded-lg hover:bg-red-50/10 transition-colors flex-shrink-0"
                            title="Remover E-mail"
                          >
                            <X className="w-4 h-4" />
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Modal de Ajuda / Tutorial do MS Teams */}
      <AnimatePresence>
        {isHelpModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/40 backdrop-blur-sm">
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="rounded-2xl sm:rounded-3xl p-5 sm:p-7 w-full max-w-2xl shadow-2xl bg-white border border-slate-200 text-slate-900 overflow-hidden flex flex-col max-h-[90vh]"
            >
              <div className="flex justify-between items-start gap-4 mb-4 pb-3 border-b border-slate-100 flex-shrink-0">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="p-2 rounded-xl bg-blue-50 text-blue-600 flex-shrink-0">
                    <HelpCircle className="w-5 h-5 sm:w-6 sm:h-6" />
                  </div>
                  <div>
                    <h3 className="text-sm sm:text-base font-black uppercase text-slate-900 tracking-tight leading-none">Como obter o Chat/Grupo ID do Teams</h3>
                    <p className="text-[9px] font-black uppercase tracking-wider text-slate-400 mt-0.5">Configure passo a passo os seus alertas operacionais no Teams</p>
                  </div>
                </div>
                <button 
                  onClick={() => { setIsHelpModalOpen(false); setTestTeamsLink(''); }} 
                  className="p-1.5 rounded-full transition-colors hover:bg-slate-100 text-slate-400 hover:text-slate-900 flex-shrink-0 cursor-pointer"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Scrollable content */}
              <div className="overflow-y-auto pr-1 space-y-4 text-xs sm:text-sm text-slate-600">
                <div className="p-3.5 bg-slate-50 border border-slate-100 rounded-xl space-y-2">
                  <p className="font-extrabold text-[10px] uppercase tracking-wide text-slate-500">Fluxo pelos Aplicativos (Teams Desktop ou Web)</p>
                  <ol className="list-decimal pl-5 space-y-1.5 font-bold text-slate-700 text-xs leading-relaxed">
                    <li>Abra o <span className="text-slate-900 font-extrabold">Microsoft Teams</span>.</li>
                    <li>Vá na aba do menu lateral <span className="text-slate-900 font-extrabold">Equipes</span> (ou nas suas conversas).</li>
                    <li>Clique nos <span className="text-slate-900 font-extrabold">três pontinhos (...)</span> localizado ao lado do nome da canal/grupo ou equipe.</li>
                    <li>Selecione <span className="text-slate-900 font-extrabold">"Obter link para a equipe" (ou "Get link to team")</span> e copie o link gerado.</li>
                    <li>Cole no campo abaixo para isolarmos o ID para você em 1 clique!</li>
                  </ol>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <div className="p-3 border border-dashed border-slate-200 rounded-xl bg-blue-50/10">
                    <h4 className="text-[9px] font-black uppercase tracking-wider text-slate-400 mb-1">Para Chat Individual / Privado comum:</h4>
                    <p className="text-[10px] leading-normal font-bold text-slate-600">
                      O link copiado costuma ser similar a:
                    </p>
                    <div className="bg-slate-100/80 text-slate-700 p-2 rounded-lg font-mono text-[9px] mt-1 break-all select-all font-semibold leading-relaxed">
                      https://teams.microsoft.com/l/chat/<span className="bg-yellow-200 text-slate-900 font-extrabold px-0.5 rounded">19:173b6b51-bde9-48e9-8125-9153dd100f4c_bb0a56bd-0e15-461e-a4d4-e33af5f669b8@unq.gbl.spaces</span>/conversations?context=...
                    </div>
                    <p className="text-[10px] mt-2 font-black text-blue-600">
                      O ID correto é a parte em destaque amarelo.
                    </p>
                  </div>

                  <div className="p-3 border border-dashed border-slate-200 rounded-xl bg-emerald-50/10">
                    <h4 className="text-[9px] font-black uppercase tracking-wider text-slate-400 mb-1">Para Chat em Grupo / Canal de Equipes:</h4>
                    <p className="text-[10px] leading-normal font-bold text-slate-600">
                      O link copiado costuma ser similar a:
                    </p>
                    <div className="bg-slate-100/80 text-slate-700 p-2 rounded-lg font-mono text-[9px] mt-1 break-all select-all font-semibold leading-relaxed">
                      https://teams.microsoft.com/l/chat/<span className="bg-yellow-200 text-slate-900 font-extrabold px-0.5 rounded">19:39fa3bb7be4c4b82a762c6af137f181c@thread.v2</span>/conversations?context=...
                    </div>
                    <p className="text-[10px] mt-2 font-black text-emerald-600">
                      O ID correto é a parte em destaque amarelo.
                    </p>
                  </div>
                </div>

                {/* Live Parser Box */}
                <div className="border-t border-slate-100 pt-3.5 space-y-2.5">
                  <div className="flex items-center gap-1.5">
                    <div className="w-1.5 h-1.5 rounded-full bg-blue-500 animate-pulse" />
                    <h4 className="text-[10px] font-black uppercase tracking-wider text-slate-800">Ferramenta: Extrator Inteligente de Link do Teams</h4>
                  </div>
                  <p className="text-[10px] leading-normal font-bold text-slate-500 uppercase tracking-widest">Cole o link gerado do Teams abaixo para isolarmos o ID sem complicações:</p>
                  
                  <textarea 
                    rows={2}
                    value={testTeamsLink}
                    onChange={(e) => setTestTeamsLink(e.target.value)}
                    placeholder="Cole o endereço completo aqui..."
                    className="w-full px-3 py-2 bg-slate-50 border border-slate-200 text-slate-900 font-bold text-xs rounded-xl outline-none focus:ring-1 focus:ring-slate-900"
                  />

                  {extractedIdResult && (
                    <div className={`p-3 rounded-xl border flex flex-col sm:flex-row sm:items-center justify-between gap-3 ${extractedIdResult.startsWith('19:') ? 'bg-emerald-50 border-emerald-100 text-emerald-800' : 'bg-red-50 border-red-100 text-red-800'}`}>
                      <div className="min-w-0 flex-grow">
                        <span className="block text-[8px] font-black uppercase tracking-widest text-slate-400 mb-0.5">ID Identificado:</span>
                        <code className="text-[10px] sm:text-[11px] font-mono font-black break-all select-all">{extractedIdResult}</code>
                      </div>
                      {extractedIdResult.startsWith('19:') && (
                        <div className="flex gap-1.5 flex-shrink-0 self-end sm:self-auto">
                          <button
                            type="button"
                            onClick={() => {
                              navigator.clipboard.writeText(extractedIdResult);
                              alert('ID copiado!');
                            }}
                            className="px-2.5 py-1.5 bg-white border border-slate-200 hover:bg-slate-50 text-slate-800 font-black text-[9px] uppercase tracking-wider rounded-lg transition-all"
                          >
                            Copiar
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setTeamsChatId(extractedIdResult);
                              setIsHelpModalOpen(false);
                              setTestTeamsLink('');
                              setTeamsMessage({ type: 'success', text: 'ID extraído e configurado no painel! Lembre-se de clicar no botão "Salvar ID" para gravar.' });
                              setTimeout(() => setTeamsMessage(null), 8000);
                            }}
                            className="px-2.5 py-1.5 bg-[#1f4e79] hover:bg-[#153552] text-white font-black text-[9px] uppercase tracking-wider rounded-lg transition-all flex items-center gap-1 shadow-sm"
                          >
                            <Save className="w-3 h-3 text-white" /> Aplicar ID
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>

              <div className="mt-4 pt-3 border-t border-slate-100 flex justify-end flex-shrink-0">
                <button
                  type="button"
                  onClick={() => { setIsHelpModalOpen(false); setTestTeamsLink(''); }}
                  className="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-800 rounded-xl font-black text-xs uppercase tracking-widest transition-all cursor-pointer shadow-sm"
                >
                  OK, FECHAR
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
