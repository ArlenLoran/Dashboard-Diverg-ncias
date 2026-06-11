import { useState, useEffect, useRef, type RefObject } from 'react';
import { createPortal } from 'react-dom';
import { 
  CheckCircle2, XCircle, ChevronLeft, ChevronRight, RefreshCcw, 
  X, Info, Download, BookOpen, ShieldCheck, Search, Filter,
  TrendingUp, TrendingDown, Activity, Settings, LayoutGrid,
  Clock, Bell, Triangle, Sparkles, Fingerprint, Users, Shield, Lock,
  Volume2, HelpCircle, Mail, Copy, Check
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import Markdown from 'react-markdown';
import html2canvas from 'html2canvas';
import { 
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, 
  ResponsiveContainer, LineChart, Line 
} from 'recharts';
import type { Section, Metric, MetricHistory } from '../types';

import { postSqlQuery } from '../services/queryService';
import { 
  ensureSharePointConfig, 
  fetchDashboardConfig, 
  saveMetricData,
  isUserAllowed,
  fetchAllowedUsers,
  addAllowedUser,
  removeAllowedUser,
  getTeamsChatId,
  getEmailAlertsEnabled,
  getTeamsAlertsEnabled,
  getAiEnabled,
  fetchAlertEmails
} from '../services/configService';
import { getCurrentSharePointUserEmail, hasSpContext } from '../services/spService';

const sendingEmailLocks = new Set<string>();

function playAlertBeep(type: 'warning' | 'critical' | 'success' = 'warning') {
  try {
    const AudioContextClass = window.AudioContext || (window as any).webkitAudioContext;
    if (!AudioContextClass) return;
    const audioCtx = new AudioContextClass();
    const osc = audioCtx.createOscillator();
    const gain = audioCtx.createGain();
    
    osc.connect(gain);
    gain.connect(audioCtx.destination);
    
    if (type === 'critical') {
      osc.type = 'sawtooth';
      osc.frequency.setValueAtTime(880, audioCtx.currentTime); // A5
      gain.gain.setValueAtTime(0.08, audioCtx.currentTime);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.35);
    } else if (type === 'success') {
      osc.type = 'sine';
      osc.frequency.setValueAtTime(523.25, audioCtx.currentTime); // C5
      gain.gain.setValueAtTime(0.08, audioCtx.currentTime);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.15);
      
      setTimeout(() => {
        try {
          const osc2 = audioCtx.createOscillator();
          const gain2 = audioCtx.createGain();
          osc2.connect(gain2);
          gain2.connect(audioCtx.destination);
          osc2.type = 'sine';
          osc2.frequency.setValueAtTime(659.25, audioCtx.currentTime); // E5
          gain2.gain.setValueAtTime(0.08, audioCtx.currentTime);
          osc2.start();
          osc2.stop(audioCtx.currentTime + 0.2);
        } catch {}
      }, 150);
    } else {
      osc.type = 'sine';
      osc.frequency.setValueAtTime(440, audioCtx.currentTime); // A4
      gain.gain.setValueAtTime(0.08, audioCtx.currentTime);
      osc.start();
      osc.stop(audioCtx.currentTime + 0.2);
    }
  } catch (e) {
    console.warn("Audio Context beep error:", e);
  }
}

function speakAlertText(text: string) {
  try {
    if ('speechSynthesis' in window) {
      window.speechSynthesis.cancel();
      const utterance = new SpeechSynthesisUtterance(text);
      utterance.lang = 'pt-BR';
      utterance.rate = 1.0;
      utterance.pitch = 1.0;
      window.speechSynthesis.speak(utterance);
    }
  } catch (e) {
    console.warn("Speech Synthesis warning:", e);
  }
}

function oklchToRgb(l: number, c: number, h: number, alpha = 1): string {
  // Convert hue to radians
  const hRad = (h * Math.PI) / 180;
  
  // OKLCH to OKLAB
  const a = c * Math.cos(hRad);
  const b = c * Math.sin(hRad);
  
  // OKLAB to LMS
  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.2914855414 * b;
  
  const l_LMS = Math.pow(Math.max(0, l_), 3);
  const m_LMS = Math.pow(Math.max(0, m_), 3);
  const s_LMS = Math.pow(Math.max(0, s_), 3);
  
  // LMS to Linear sRGB
  const rL = +4.0767416621 * l_LMS - 3.3077115913 * m_LMS + 0.2309699292 * s_LMS;
  const gL = -1.2684380046 * l_LMS + 2.6097574011 * m_LMS - 0.3413193965 * s_LMS;
  const bL = -0.0041960863 * l_LMS - 0.7034186147 * m_LMS + 1.7076147010 * s_LMS;
  
  // Linear sRGB to sRGB
  const toSRGB = (val: number) => {
    return val <= 0.0031308 ? 12.92 * val : 1.055 * Math.pow(val, 1 / 2.4) - 0.055;
  };
  
  const outR = Math.min(255, Math.max(0, Math.round(toSRGB(rL) * 255)));
  const outG = Math.min(255, Math.max(0, Math.round(toSRGB(gL) * 255)));
  const outB = Math.min(255, Math.max(0, Math.round(toSRGB(bL) * 255)));
  
  if (alpha < 1) {
    return `rgba(${outR}, ${outG}, ${outB}, ${alpha})`;
  }
  return `rgb(${outR}, ${outG}, ${outB})`;
}

function oklabToRgb(l: number, a: number, b: number, alpha = 1): string {
  // OKLAB to LMS
  const l_ = l + 0.3963377774 * a + 0.2158037573 * b;
  const m_ = l - 0.1055613458 * a - 0.0638541728 * b;
  const s_ = l - 0.0894841775 * a - 1.2914855414 * b;
  
  const l_LMS = Math.pow(Math.max(0, l_), 3);
  const m_LMS = Math.pow(Math.max(0, m_), 3);
  const s_LMS = Math.pow(Math.max(0, s_), 3);
  
  // LMS to Linear sRGB
  const rL = +4.0767416621 * l_LMS - 3.3077115913 * m_LMS + 0.2309699292 * s_LMS;
  const gL = -1.2684380046 * l_LMS + 2.6097574011 * m_LMS - 0.3413193965 * s_LMS;
  const bL = -0.0041960863 * l_LMS - 0.7034186147 * m_LMS + 1.7076147010 * s_LMS;
  
  // Linear sRGB to sRGB
  const toSRGB = (val: number) => {
    return val <= 0.0031308 ? 12.92 * val : 1.055 * Math.pow(val, 1 / 2.4) - 0.055;
  };
  
  const outR = Math.min(255, Math.max(0, Math.round(toSRGB(rL) * 255)));
  const outG = Math.min(255, Math.max(0, Math.round(toSRGB(gL) * 255)));
  const outB = Math.min(255, Math.max(0, Math.round(toSRGB(bL) * 255)));
  
  if (alpha < 1) {
    return `rgba(${outR}, ${outG}, ${outB}, ${alpha})`;
  }
  return `rgb(${outR}, ${outG}, ${outB})`;
}

function replaceOklchOklabValues(cssText: string): string {
  if (!cssText || typeof cssText !== 'string') return cssText;
  
  // Match oklch(...) ou oklab(...)
  const colorRegex = /(oklch|oklab)\(([^)]+)\)/gi;
  
  return cssText.replace(colorRegex, (match, type, content) => {
    try {
      if (content.includes('var(')) {
        return 'rgb(71, 85, 105)';
      }
      
      const sanitized = content.replace(/\//g, ' ').replace(/,/g, ' ').trim();
      const tokens = sanitized.split(/\s+/).filter(Boolean);
      
      if (tokens.length < 3) {
        return 'rgb(71, 85, 105)';
      }
      
      let val1 = parseFloat(tokens[0]);
      if (tokens[0].endsWith('%')) {
        val1 = parseFloat(tokens[0]) / 100;
      }
      let val2 = parseFloat(tokens[1]);
      if (tokens[1].endsWith('%')) {
        val2 = parseFloat(tokens[1]) / 100;
      }
      let val3 = parseFloat(tokens[2]);
      if (tokens[2].endsWith('%')) {
        val3 = parseFloat(tokens[2]) / 100;
      }
      
      let alpha = 1;
      if (tokens.length >= 4) {
        alpha = parseFloat(tokens[3]);
        if (tokens[3].endsWith('%')) {
          alpha = parseFloat(tokens[3]) / 100;
        }
      }
      
      if (isNaN(val1) || isNaN(val2) || isNaN(val3)) {
        return 'rgb(71, 85, 105)';
      }
      
      if (type.toLowerCase() === 'oklch') {
        return oklchToRgb(val1, val2, val3, alpha);
      } else {
        return oklabToRgb(val1, val2, val3, alpha);
      }
    } catch (e) {
      return 'rgb(71, 85, 105)';
    }
  });
}

const withCleanedEnvironment = async <T,>(actionFn: () => Promise<T>): Promise<T> => {
  const originalGetComputedStyle = window.getComputedStyle;
  const originalCreateElement = document.createElement;
  const originalFetch = window.fetch;
  const originalXHR = window.XMLHttpRequest;
  
  const originalProtoGetPropertyValue = CSSStyleDeclaration.prototype.getPropertyValue;
  
  // 1. Intercept CSSStyleDeclaration.prototype.getPropertyValue
  CSSStyleDeclaration.prototype.getPropertyValue = function(property: string) {
    const val = originalProtoGetPropertyValue.call(this, property);
    if (val && (val.includes('oklch') || val.includes('oklab'))) {
      return replaceOklchOklabValues(val);
    }
    return val;
  };

  // 1a. Intercept all properties with getters on standard CSSStyleDeclaration prototype
  const descriptors = Object.getOwnPropertyDescriptors(CSSStyleDeclaration.prototype);
  const restoredDescriptors: Record<string, PropertyDescriptor> = {};

  for (const key of Object.keys(descriptors)) {
    const desc = descriptors[key];
    if (desc && desc.get && typeof desc.get === 'function') {
      restoredDescriptors[key] = desc;
      try {
        Object.defineProperty(CSSStyleDeclaration.prototype, key, {
          get() {
            const val = desc.get!.call(this);
            if (typeof val === 'string' && (val.includes('oklch') || val.includes('oklab'))) {
              return replaceOklchOklabValues(val);
            }
            return val;
          },
          configurable: true,
          enumerable: desc.enumerable
        });
      } catch (e) {
        // Can't redefine some properties
      }
    }
  }
  
  // Helper to wrap getComputedStyle on a given window
  const wrapWinGetComputedStyle = (win: any) => {
    const orig = win.getComputedStyle;
    win.getComputedStyle = function(el: Element, pseudo?: string) {
      const style = orig.call(win, el, pseudo);
      return new Proxy(style, {
        get(target, prop) {
          if (prop === 'getPropertyValue') {
            return function(name: string) {
              const val = target.getPropertyValue(name);
              if (val && (val.includes('oklch') || val.includes('oklab'))) {
                return replaceOklchOklabValues(val);
              }
              return val;
            };
          }
          const val = Reflect.get(target, prop);
          if (typeof val === 'string' && (val.includes('oklch') || val.includes('oklab'))) {
            return replaceOklchOklabValues(val);
          }
          if (typeof val === 'function') {
            return val.bind(target);
          }
          return val;
        }
      });
    };
  };

  // 2. Intercept getComputedStyle on the main window
  wrapWinGetComputedStyle(window);

  // 3. Intercept iframe creation to inject our wrap into iframe windows
  document.createElement = function(tagName: string, options?: ElementCreationOptions) {
    const el = originalCreateElement.call(document, tagName, options);
    if (tagName.toLowerCase() === 'iframe') {
      Object.defineProperty(el, 'contentWindow', {
        get() {
          const iframeWin = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'contentWindow')?.get?.call(this);
          if (iframeWin && !iframeWin.__getComputedStyleOverridden__) {
            iframeWin.__getComputedStyleOverridden__ = true;
            wrapWinGetComputedStyle(iframeWin);
            try {
              if (iframeWin.CSSStyleDeclaration) {
                const innerProto = iframeWin.CSSStyleDeclaration.prototype;
                const innerOrigGetProp = innerProto.getPropertyValue;
                innerProto.getPropertyValue = function(propName: string) {
                  const val = innerOrigGetProp.call(this, propName);
                  if (val && (val.includes('oklch') || val.includes('oklab'))) {
                    return replaceOklchOklabValues(val);
                  }
                  return val;
                };
              }
            } catch (e) {}
          }
          return iframeWin;
        },
        configurable: true
      });
    }
    return el;
  } as any;

  // 4. Intercept fetch to clean any loaded CSS files
  window.fetch = async function(input: RequestInfo | URL, init?: RequestInit) {
    const response = await originalFetch.call(window, input, init);
    const url = typeof input === 'string' ? input : (input instanceof Request ? input.url : '');
    if (url.includes('.css') || (init && init.headers && String((init.headers as any)['Accept'] || '').includes('text/css'))) {
      try {
        const text = await response.text();
        const cleanedText = replaceOklchOklabValues(text);
        return new Response(cleanedText, {
          status: response.status,
          statusText: response.statusText,
          headers: response.headers
        });
      } catch (err) {
        // Fallback to original
      }
    }
    return response;
  };

  // 5. Intercept XMLHttpRequest to clean loaded CSS
  window.XMLHttpRequest = function() {
    const xhr = new originalXHR();
    const originalOpen = xhr.open;
    let isCss = false;
    xhr.open = function(method: string, url: string | URL, ...args: any[]) {
      if (typeof url === 'string' && url.includes('.css')) {
        isCss = true;
      }
      return originalOpen.call(xhr, method, url, ...args);
    } as any;

    Object.defineProperty(xhr, 'responseText', {
      get() {
        const raw = xhr.response;
        if (isCss && typeof raw === 'string') {
          return replaceOklchOklabValues(raw);
        }
        return raw;
      },
      configurable: true
    });
    return xhr;
  } as any;

  try {
    return await actionFn();
  } finally {
    // Restore everything perfectly!
    window.getComputedStyle = originalGetComputedStyle;
    document.createElement = originalCreateElement;
    window.fetch = originalFetch;
    window.XMLHttpRequest = originalXHR;
    CSSStyleDeclaration.prototype.getPropertyValue = originalProtoGetPropertyValue;
    
    for (const key of Object.keys(restoredDescriptors)) {
      try {
        Object.defineProperty(CSSStyleDeclaration.prototype, key, restoredDescriptors[key]);
      } catch (e) {}
    }
  }
};

function copyInputStates(original: HTMLElement, clone: HTMLElement) {
  const originalInputs = Array.from(original.querySelectorAll('input, select, textarea')) as (HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement)[];
  const cloneInputs = Array.from(clone.querySelectorAll('input, select, textarea')) as (HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement)[];
  
  for (let i = 0; i < originalInputs.length; i++) {
    const orig = originalInputs[i];
    const cln = cloneInputs[i];
    if (!orig || !cln) continue;
    
    if (orig.tagName.toLowerCase() === 'input') {
      const type = (orig as HTMLInputElement).type;
      if (type === 'checkbox' || type === 'radio') {
        (cln as HTMLInputElement).checked = (orig as HTMLInputElement).checked;
      } else {
        (cln as HTMLInputElement).value = (orig as HTMLInputElement).value;
      }
    } else if (orig.tagName.toLowerCase() === 'textarea') {
      (cln as HTMLTextAreaElement).value = (orig as HTMLTextAreaElement).value;
    } else if (orig.tagName.toLowerCase() === 'select') {
      (cln as HTMLSelectElement).selectedIndex = (orig as HTMLSelectElement).selectedIndex;
    }
  }
}

function copyCanvases(original: HTMLElement, clone: HTMLElement) {
  const originalCanvases = Array.from(original.querySelectorAll('canvas')) as HTMLCanvasElement[];
  const cloneCanvases = Array.from(clone.querySelectorAll('canvas')) as HTMLCanvasElement[];
  
  for (let i = 0; i < originalCanvases.length; i++) {
    const orig = originalCanvases[i];
    const cln = cloneCanvases[i];
    if (!orig || !cln) continue;
    
    try {
      const ctx = cln.getContext('2d');
      if (ctx) {
        ctx.drawImage(orig, 0, 0);
      }
    } catch (e) {
      // ignore
    }
  }
}

function copyStylesRecursive(original: Element, clone: Element) {
  if (!(original instanceof HTMLElement || original instanceof SVGElement) ||
      !(clone instanceof HTMLElement || clone instanceof SVGElement)) {
    return;
  }
  const style = window.getComputedStyle(original);
  
  const propertiesToCopy = [
    'color', 'background-color', 'background',
    'border-color', 'border-top-color', 'border-right-color', 'border-bottom-color', 'border-left-color',
    'box-shadow', 'fill', 'stroke', 'outline-color', 'outline',
    'background-image'
  ];
  
  for (const prop of propertiesToCopy) {
    try {
      let val = style.getPropertyValue(prop);
      if (val) {
        if (val.includes('oklch') || val.includes('oklab')) {
          val = replaceOklchOklabValues(val);
        }
        clone.style.setProperty(prop, val, 'important');
      }
    } catch (e) {
      // Safe skip
    }
  }
  
  const inlineStyle = clone.getAttribute('style');
  if (inlineStyle && (inlineStyle.includes('oklch') || inlineStyle.includes('oklab'))) {
    clone.setAttribute('style', replaceOklchOklabValues(inlineStyle));
  }

  const origChildren = Array.from(original.children);
  const cloneChildren = Array.from(clone.children);
  const minLength = Math.min(origChildren.length, cloneChildren.length);
  for (let i = 0; i < minLength; i++) {
    copyStylesRecursive(origChildren[i], cloneChildren[i]);
  }
}

interface SafeCloneResult {
  clone: HTMLElement;
  cleanup: () => void;
}

function createHtml2CanvasSafeClone(targetElement: HTMLElement, isWarRoom: boolean): SafeCloneResult {
  const rect = targetElement.getBoundingClientRect();
  const width = rect.width || targetElement.scrollWidth || 1200;
  
  const clone = targetElement.cloneNode(true) as HTMLElement;

  // 1. Copiar estados e estilos de forma recursiva na árvore 1-para-1 perfeitamente intacta
  try {
    copyInputStates(targetElement, clone);
    copyCanvases(targetElement, clone);
  } catch (e) {
    console.warn("[Teams/Clone] Error copying element states:", e);
  }
  
  try {
    copyStylesRecursive(targetElement, clone);
  } catch (e) {
    console.warn("[Teams/Clone] Error copying and cleaning styles recursively:", e);
  }

  // 2. Agora, com os estilos aplicados de forma precisa, remover elementos indesejados do clone
  const headerToIgnore = clone.querySelector('#dashboard-header-print-ignore');
  if (headerToIgnore && headerToIgnore.parentNode) {
    headerToIgnore.parentNode.removeChild(headerToIgnore);
  }
  const kpiBlockToIgnore = clone.querySelector('#dashboard-kpi-block-print-ignore');
  if (kpiBlockToIgnore && kpiBlockToIgnore.parentNode) {
    kpiBlockToIgnore.parentNode.removeChild(kpiBlockToIgnore);
  }
  
  const wrapper = document.createElement('div');
  wrapper.id = 'html2canvas-safe-clone-wrapper';
  wrapper.style.position = 'fixed';
  wrapper.style.left = '-15000px';
  wrapper.style.top = '0px';
  wrapper.style.width = `${width}px`;
  wrapper.style.height = 'auto';
  wrapper.style.overflow = 'hidden';
  wrapper.style.zIndex = '-99999';
  wrapper.style.pointerEvents = 'none';
  wrapper.style.backgroundColor = isWarRoom ? '#050510' : '#f8fafc';
  
  clone.style.width = `${width}px`;
  clone.style.height = 'auto';
  clone.style.margin = '0px';
  clone.style.padding = '24px';
  clone.style.boxSizing = 'border-box';
  
  wrapper.appendChild(clone);
  document.body.appendChild(wrapper);
  
  return {
    clone: wrapper,
    cleanup: () => {
      if (wrapper && wrapper.parentNode) {
        wrapper.parentNode.removeChild(wrapper);
      }
    }
  };
}

interface MetricCardProps {
  metric: Metric;
  onClick?: (metric: Metric) => void;
  onRefresh?: (metric: Metric) => void;
  isWarRoom?: boolean;
}

function useScrollIndicator(ref: RefObject<HTMLDivElement | null>) {
  const [scrollInfo, setScrollInfo] = useState({ percentage: 0, ratio: 0, isScrollable: false });

  const updateScroll = () => {
    if (ref.current) {
      const { scrollLeft, scrollWidth, clientWidth } = ref.current;
      const isScrollable = scrollWidth > clientWidth + 1;
      const percentage = isScrollable ? (scrollLeft / (scrollWidth - clientWidth)) * 100 : 0;
      const ratio = isScrollable ? clientWidth / scrollWidth : 1;
      setScrollInfo({ percentage, ratio, isScrollable });
    }
  };

  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    updateScroll();
    const observer = new ResizeObserver(updateScroll);
    observer.observe(el);
    el.addEventListener('scroll', updateScroll);
    return () => {
      observer.disconnect();
      el.removeEventListener('scroll', updateScroll);
    };
  }, [ref]);

  return scrollInfo;
}

function Countdown({ metric, onRefresh, hideUI, className }: { metric: Metric, onRefresh?: (m: Metric) => void, hideUI?: boolean, className?: string }) {
  const [timeLeft, setTimeLeft] = useState<number | null>(null);
  const isRefreshingRef = useRef(false);
  const onRefreshRef = useRef(onRefresh);

  useEffect(() => {
    onRefreshRef.current = onRefresh;
  }, [onRefresh]);

  useEffect(() => {
    if (!metric || !metric.refreshInterval || !onRefreshRef.current) return;
    
    isRefreshingRef.current = false;

    const calculate = () => {
      if (!metric.lastUpdateAt) return 0;
      const now = new Date();
      const last = new Date(metric.lastUpdateAt);
      if (isNaN(last.getTime())) return 0;
      const next = new Date(last.getTime() + metric.refreshInterval! * 60000);
      const diff = Math.max(0, Math.floor((next.getTime() - now.getTime()) / 1000));
      return diff;
    };

    const initial = calculate();
    setTimeLeft(initial);
    
    // If we're already at zero, trigger refresh once
    if (initial <= 0 && !isRefreshingRef.current) {
      isRefreshingRef.current = true;
      if (onRefreshRef.current) {
        onRefreshRef.current(metric);
      }
    }

    const timer = setInterval(() => {
      const remaining = calculate();
      setTimeLeft(remaining);
      if (remaining <= 0 && !isRefreshingRef.current) {
        isRefreshingRef.current = true;
        if (onRefreshRef.current) {
          onRefreshRef.current(metric);
        }
      } else if (remaining > 0) {
        // Reset the flag if we have time again (meaning data updated)
        isRefreshingRef.current = false;
      }
    }, 1000);
    
    return () => {
      clearInterval(timer);
    };
  }, [metric.lastUpdateAt, metric.refreshInterval, metric.id]);

  if (hideUI) return null;
  if (timeLeft === null || !metric.refreshInterval) return null;
  const mins = Math.floor(timeLeft / 60);
  const secs = timeLeft % 60;
  const isUrgent = timeLeft < 30;

  return (
    <div className={`${className || 'flex items-center gap-1 text-[8px] font-black transition-colors duration-500'} ${isUrgent ? 'text-red-500 animate-pulse' : 'text-slate-400'}`}>
      <RefreshCcw className={`w-2 h-2 ${isUrgent ? 'animate-spin' : ''}`} />
      <span>{mins}:{secs < 10 ? '0' : ''}{secs}s</span>
    </div>
  );
}

function MiniSparkline({ data, color }: { data: MetricHistory[], color: string }) {
  const chartData = data.map((item, i) => ({ value: item.value, index: i }));
  return (
    <div className="w-full h-8 mt-2 overflow-hidden opacity-50 group-hover:opacity-100 transition-opacity duration-500">
      <ResponsiveContainer width="100%" height="100%">
        <AreaChart data={chartData}>
          <defs>
            <linearGradient id={`gradient-${color}`} x1="0" y1="0" x2="0" y2="1">
              <stop offset="5%" stopColor={color} stopOpacity={0.3}/>
              <stop offset="95%" stopColor={color} stopOpacity={0}/>
            </linearGradient>
          </defs>
          <Area 
            type="monotone" 
            dataKey="value" 
            stroke={color} 
            strokeWidth={1.5} 
            fillOpacity={1} 
            fill={`url(#gradient-${color})`} 
            isAnimationActive={false}
          />
        </AreaChart>
      </ResponsiveContainer>
    </div>
  );
}

function MetricCard({ metric, onClick, onRefresh, isWarRoom }: MetricCardProps) {
  return (
    <motion.div
      layout
      initial={{ opacity: 0, scale: 0.95 }}
      animate={{ opacity: 1, scale: 1 }}
      onClick={() => onClick?.(metric)}
      whileHover={{ 
        y: -4, 
        boxShadow: "0 20px 25px -5px rgb(0 0 0 / 0.1), 0 8px 10px -6px rgb(0 0 0 / 0.1)",
        transition: { duration: 0.2 } 
      }}
      className={`rounded-xl shadow-sm p-3 w-[160px] h-[240px] flex-shrink-0 flex flex-col items-center justify-between border transition-all duration-500 cursor-pointer relative overflow-hidden ${
        isWarRoom 
        ? 'bg-[#0b0e1a] border-indigo-900/30 shadow-slate-950/50 hover:border-indigo-500/50' 
        : 'bg-white border-slate-100 hover:border-slate-300'
      }`}
      id={`card-${metric.id}`}
    >
      <header className="w-full text-center">
        <h3 className={`text-[8px] font-bold uppercase tracking-wide truncate py-1 transition-colors duration-500 px-1 ${isWarRoom ? 'text-slate-300' : 'text-slate-700'}`}>
          {metric.title}
        </h3>
        <div className={`h-px w-3/4 mx-auto transition-colors duration-500 ${isWarRoom ? 'bg-slate-800' : 'bg-slate-200'}`} />
      </header>
      <div className="flex flex-col items-center gap-1.5 py-0.5">
        <span className={`text-lg font-black italic tracking-tighter tabular-nums transition-colors duration-500 ${isWarRoom ? 'text-white text-xl' : 'text-slate-900'}`}>
          {metric.value}
        </span>
        <div className={`relative group/icon flex items-center justify-center min-h-[52px] rounded-full transition-colors duration-500 ${isWarRoom ? 'bg-slate-800/50' : 'bg-transparent'}`}>
          {metric.status === 'ok' ? (
            <motion.div whileHover={{ scale: 1.1 }} transition={{ type: "spring", stiffness: 400, damping: 10 }}>
              <CheckCircle2 className="w-12 h-12 text-emerald-500" strokeWidth={2.5} />
            </motion.div>
          ) : (
            <motion.div
              animate={{ scale: [1, isWarRoom ? 1.2 : 1.15, 1] }}
              transition={{ repeat: Infinity, duration: isWarRoom ? 0.6 : 0.8, ease: "easeInOut" }}
              style={{ filter: isWarRoom ? "drop-shadow(0 0 8px rgba(255, 0, 0, 0.4))" : "drop-shadow(0 0 6px rgba(204, 0, 0, 0.25))" }}
              whileHover={{ scale: 1.25 }}
            >
              <XCircle className={`w-12 h-12 ${isWarRoom ? 'text-red-500' : 'text-brand-red'}`} strokeWidth={2.5} />
            </motion.div>
          )}
        </div>
      </div>
      <footer className="w-full text-center">
        <div className={`h-px w-3/4 mx-auto mb-1 transition-colors duration-500 ${isWarRoom ? 'bg-slate-800' : 'bg-slate-200'}`} />
        {metric.history && metric.history.length > 2 && (
          <MiniSparkline data={metric.history.map((val, idx) => ({ value: val, timestamp: String(idx) }))} color={metric.status === 'ok' ? '#10b981' : '#ef4444'} />
        )}
        <div className="flex items-center justify-center gap-1 mt-1.5">
          <div className="flex items-center gap-1">
            <Clock className={`w-2 h-2 ${isWarRoom ? 'text-slate-600' : 'text-slate-300'}`} />
            <span className={`text-[8px] font-bold tracking-tighter transition-colors duration-500 ${isWarRoom ? 'text-slate-500' : 'text-slate-400'}`}>
              {metric.lastUpdate}
            </span>
          </div>
          <Countdown metric={metric} onRefresh={onRefresh} hideUI />
        </div>
      </footer>
      {isWarRoom && metric.status === 'error' && <div className="absolute inset-x-0 bottom-0 h-1 bg-red-500 animate-pulse" />}
    </motion.div>
  );
}

function SectionContainer({ section, onCardClick, onCardRefresh, isWarRoom }: { section: Section, onCardClick: (metric: Metric) => void, onCardRefresh: (metric: Metric) => void, isWarRoom: boolean }) {
  const scrollRef = useRef<HTMLDivElement>(null);
  const { percentage, ratio, isScrollable } = useScrollIndicator(scrollRef);
  const scrollInterval = useRef<NodeJS.Timeout | null>(null);

  useEffect(() => {
    if (isWarRoom && isScrollable) {
      scrollInterval.current = setInterval(() => {
        if (scrollRef.current) {
          const { scrollLeft, scrollWidth, clientWidth } = scrollRef.current;
          if (scrollLeft + clientWidth >= scrollWidth - 5) {
            scrollRef.current.scrollTo({ left: 0, behavior: 'smooth' });
          } else {
            scrollRef.current.scrollBy({ left: 300, behavior: 'smooth' });
          }
        }
      }, 6000);
    } else {
      if (scrollInterval.current) clearInterval(scrollInterval.current);
    }
    return () => { if (scrollInterval.current) clearInterval(scrollInterval.current); };
  }, [isWarRoom, isScrollable]);

  const scroll = (direction: 'left' | 'right') => {
    if (scrollRef.current) {
      const amount = direction === 'left' ? -300 : 300;
      scrollRef.current.scrollBy({ left: amount, behavior: 'smooth' });
    }
  };

  return (
    <section className={`flex flex-col gap-3 transition-all duration-500 ${isWarRoom ? 'scale-[1.02]' : ''}`} id={`section-${section.title.replace(/\s/g, '-')}`}>
      <div className={`py-4 px-8 rounded-2xl shadow-lg inline-flex items-center w-full transition-colors duration-500 transform -skew-x-6 ${isWarRoom ? 'bg-[#0b0e1a] border border-indigo-900/30' : 'bg-brand-red text-white'}`}>
        <h2 className={`text-xl font-black tracking-tighter uppercase italic skew-x-6 ${isWarRoom ? 'text-brand-red drop-shadow-[0_0_8px_rgba(204,0,0,0.3)]' : 'text-white'}`}>
          {section.title}
        </h2>
      </div>
      <div className="relative group flex flex-col w-full">
        <div ref={scrollRef} className="overflow-x-auto pb-4 scrollbar-hide flex gap-6 scroll-smooth px-4 w-full justify-start">
          {section.metrics.map((metric) => (
            <MetricCard key={metric.id} metric={metric} onClick={onCardClick} onRefresh={onCardRefresh} isWarRoom={isWarRoom} />
          ))}
        </div>
        <AnimatePresence>
          {isScrollable && (
            <motion.div initial={{ opacity: 0, y: 5 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: 5 }} className="flex items-center justify-between mt-2 px-4 w-full">
              <button onClick={() => scroll('left')} className={`focus:outline-none hover:scale-110 transition-all active:scale-95 ${isWarRoom ? 'text-slate-500 hover:text-white' : 'text-brand-yellow'}`}>
                <ChevronLeft className="w-5 h-5" strokeWidth={isWarRoom ? 2 : 4} />
              </button>
              <div className={`mx-4 h-1.5 flex-grow rounded-full relative overflow-hidden shadow-inner transition-colors duration-500 ${isWarRoom ? 'bg-slate-900' : 'bg-brand-yellow/10'}`}>
                <motion.div 
                  className={`absolute top-0 left-0 h-full rounded-full transition-colors duration-500 ${isWarRoom ? 'bg-brand-red shadow-[0_0_8px_rgba(239,68,68,0.6)]' : 'bg-brand-yellow'}`}
                  style={{ width: `${ratio * 100}%`, left: `${percentage * (1 - ratio)}%` }}
                  transition={{ type: 'spring', bounce: 0, duration: 0.1 }}
                />
              </div>
              <button onClick={() => scroll('right')} className={`focus:outline-none hover:scale-110 transition-all active:scale-95 ${isWarRoom ? 'text-slate-500 hover:text-white' : 'text-brand-yellow'}`}>
                <ChevronRight className="w-5 h-5" strokeWidth={isWarRoom ? 2 : 4} />
              </button>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
    </section>
  );
}

function DivergenceModal({ metric, onClose, onRefresh, enableAI = true }: { metric: Metric, onClose: () => void, onRefresh?: (m: Metric) => void, enableAI?: boolean }) {
  const [activeTab, setActiveTab] = useState<'table' | 'objective' | 'rules' | 'trend' | 'ai'>('table');
  const [searchQuery, setSearchQuery] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [aiAnalysis, setAiAnalysis] = useState<string | null>(null);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const itemsPerPage = 10;

  // If AI is disabled but we somehow started with activeTab === 'ai', reset to table
  useEffect(() => {
    if (!enableAI && activeTab === 'ai') {
      setActiveTab('table');
    }
  }, [enableAI, activeTab]);
  
  const runAiAnalysis = async () => {
    if (isAnalyzing) return;
    setIsAnalyzing(true);
    setActiveTab('ai');
    try {
      const resp = await fetch('/api/ai/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          metricTitle: metric.title,
          objective: metric.objective,
          rules: metric.rules,
          data: metric.details,
          history: metric.history
        })
      });
      const data = await resp.json();
      if (data.error) throw new Error(data.error);
      setAiAnalysis(data.analysis);
    } catch (err: any) {
      setAiAnalysis(`### ❌ Erro na análise\nNão foi possível processar a análise no momento. Detalhes: ${err.message}`);
    } finally {
      setIsAnalyzing(false);
    }
  };
  
  if (!metric) return null;

  const filteredDetails = (metric.details || []).filter(item => {
    const query = searchQuery.toLowerCase();
    if (metric.isDynamic) {
      return Object.values(item).some(val => String(val ?? '').toLowerCase().includes(query));
    }
    return (
      (item.posicao || '').toLowerCase().includes(query) ||
      (item.item || '').toLowerCase().includes(query) ||
      (item.validade || '').toLowerCase().includes(query) ||
      (item.lote || '').toLowerCase().includes(query) ||
      (item.quantidade || 0).toString().includes(query) ||
      (item.motivo || '').toLowerCase().includes(query)
    );
  });

  const totalPages = Math.ceil(filteredDetails.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const paginatedDetails = filteredDetails.slice(startIndex, startIndex + itemsPerPage);
  const dynamicColumns = metric.isDynamic && metric.details && metric.details.length > 0 ? Object.keys(metric.details[0]) : [];

  const downloadExcel = () => {
    if (!metric.details || metric.details.length === 0) return;
    let worksheet = metric.isDynamic ? XLSX.utils.json_to_sheet(metric.details) : XLSX.utils.json_to_sheet(metric.details.map(d => ({
        'Posição': d.posicao, 'Item': d.item, 'Validade': d.validade, 'Lote': d.lote, 'Quantidade': d.quantidade, 'Motivo': d.motivo
    })));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Divergências");
    XLSX.writeFile(workbook, `Divergencias_${metric.title.replace(/\s/g, '_')}.xlsx`);
  };

  return (
    <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm" onClick={onClose}>
      <motion.div initial={{ opacity: 0, scale: 0.9, y: 20 }} animate={{ opacity: 1, scale: 1, y: 0 }} exit={{ opacity: 0, scale: 0.9, y: 20 }} className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden" onClick={(e) => e.stopPropagation()}>
        <header className="px-6 py-4 border-b border-slate-100 flex flex-col sm:flex-row justify-between items-start sm:items-center bg-slate-50/50 gap-4">
          <div className="flex items-center gap-3">
            <div className={`w-2 h-8 rounded-full ${metric.status === 'error' ? 'bg-brand-red' : 'bg-emerald-500'}`} />
            <div>
              <h2 className="text-lg font-black text-slate-900 uppercase tracking-tight italic">Detalhes: <span className="text-brand-red">{metric.title}</span></h2>
              <div className="flex items-center gap-2">
                <p className="text-[10px] font-bold text-slate-500 uppercase">Variações Identificadas - {metric.lastUpdate}</p>
                {metric.refreshInterval && (
                  <>
                    <div className="w-px h-3 bg-slate-200" />
                    <div className="flex items-center gap-1">
                      <span className="text-[9px] font-black text-slate-400 uppercase tracking-tighter">Próxima atualização em:</span>
                      <Countdown metric={metric} onRefresh={onRefresh} className="flex items-center gap-1 text-[10px] font-black" />
                    </div>
                  </>
                )}
              </div>
            </div>
          </div>
          <div className="flex items-center gap-2 self-end sm:self-auto">
            <div className="flex bg-slate-200 p-1 rounded-lg">
              <button onClick={() => setActiveTab('table')} className={`p-2 rounded-md transition-all ${activeTab === 'table' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}><Info className="w-5 h-5" /></button>
              <button onClick={() => setActiveTab('objective')} className={`p-2 rounded-md transition-all ${activeTab === 'objective' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}><BookOpen className="w-5 h-5" /></button>
              <button onClick={() => setActiveTab('rules')} className={`p-2 rounded-md transition-all ${activeTab === 'rules' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}><ShieldCheck className="w-5 h-5" /></button>
              <button onClick={() => setActiveTab('trend')} className={`p-2 rounded-md transition-all ${activeTab === 'trend' ? 'bg-white shadow-sm text-slate-900' : 'text-slate-500 hover:text-slate-700'}`}><TrendingUp className="w-5 h-5" /></button>
              {enableAI && (
                <button onClick={runAiAnalysis} className={`p-2 rounded-md transition-all ${activeTab === 'ai' ? 'bg-indigo-600 text-white shadow-sm' : 'text-indigo-500 hover:text-indigo-700'}`} disabled={isAnalyzing}><Sparkles className={`w-5 h-5 ${isAnalyzing ? 'animate-pulse' : ''}`} /></button>
              )}
            </div>
            <button onClick={downloadExcel} className="p-2 bg-emerald-500 text-white rounded-lg hover:bg-emerald-600 transition-all shadow-sm active:scale-95 disabled:opacity-50" disabled={!metric.details || metric.details.length === 0}><Download className="w-5 h-5" /></button>
            <button onClick={onClose} className="p-2 hover:bg-slate-200 rounded-full transition-colors text-slate-400 hover:text-slate-600"><X className="w-6 h-6" /></button>
          </div>
        </header>

        <div className="flex-grow overflow-auto p-6">
          <AnimatePresence mode="wait">
            {activeTab === 'table' && (
              <motion.div key="table" initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 10 }} transition={{ duration: 0.2 }} className="space-y-4">
                {metric.details && metric.details.length > 0 ? (
                  <>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-400" />
                      <input type="text" placeholder="Buscar em qualquer coluna..." value={searchQuery} onChange={(e) => { setSearchQuery(e.target.value); setCurrentPage(1); }} className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:ring-2 focus:ring-slate-200 focus:border-slate-400 outline-none transition-all font-medium" />
                    </div>
                    <div className="border border-slate-200 rounded-xl overflow-hidden shadow-sm overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-slate-50 border-b border-slate-200">
                            {metric.isDynamic ? dynamicColumns.map(col => <th key={col} className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">{col}</th>) : (
                                <>
                                  <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">Posição</th>
                                  <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">Item / Descrição</th>
                                  <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider text-center">Validade</th>
                                  <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider text-center">Lote</th>
                                  <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider text-center">Qtd.</th>
                                  <th className="px-4 py-3 text-xs font-black text-slate-600 uppercase tracking-wider">Motivo Divergência</th>
                                </>
                            )}
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {paginatedDetails.length > 0 ? paginatedDetails.map((detail, idx) => (
                              <tr key={detail.id || idx} className="hover:bg-slate-50 transition-colors group">
                                {metric.isDynamic ? dynamicColumns.map(col => <td key={col} className="px-4 py-4 text-sm font-medium text-slate-700">{String(detail[col] ?? '')}</td>) : (
                                    <>
                                      <td className="px-4 py-4 text-sm font-bold text-slate-900 font-mono">{detail.posicao}</td>
                                      <td className="px-4 py-4 text-sm font-medium text-slate-700">{detail.item}</td>
                                      <td className="px-4 py-4 text-sm text-slate-600 text-center font-mono">{detail.validade}</td>
                                      <td className="px-4 py-4 text-sm text-slate-600 text-center font-mono">{detail.lote}</td>
                                      <td className="px-4 py-4 text-sm font-black text-brand-red text-center tabular-nums">{detail.quantidade}</td>
                                      <td className="px-4 py-4 text-sm"><span className="px-2 py-1 bg-red-50 text-red-700 rounded-lg font-bold text-[10px] border border-red-100">{detail.motivo}</span></td>
                                    </>
                                )}
                              </tr>
                          )) : <tr><td colSpan={6} className="px-4 py-10 text-center font-medium text-slate-500 italic">Nenhum resultado encontrado para "{searchQuery}"</td></tr>}
                        </tbody>
                      </table>
                    </div>
                    {totalPages > 1 && (
                      <div className="flex items-center justify-between px-2 pt-2">
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Exibindo {startIndex + 1}-{Math.min(startIndex + itemsPerPage, filteredDetails.length)} de {filteredDetails.length} itens</span>
                        <div className="flex gap-1">
                          <button onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))} disabled={currentPage === 1} className="p-2 hover:bg-slate-100 rounded-lg disabled:opacity-30 transition-colors"><ChevronLeft className="w-5 h-5" /></button>
                          <div className="flex items-center gap-1 px-2">
                            {Array.from({ length: totalPages }, (_, i) => i + 1).map((page) => (
                              <button key={page} onClick={() => setCurrentPage(page)} className={`w-8 h-8 rounded-lg text-xs font-bold transition-all ${currentPage === page ? 'bg-slate-900 text-white shadow-md' : 'text-slate-400 hover:bg-slate-100'}`}>{page}</button>
                            ))}
                          </div>
                          <button onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))} disabled={currentPage === totalPages} className="p-2 hover:bg-slate-100 rounded-lg disabled:opacity-30 transition-colors"><ChevronRight className="w-5 h-5" /></button>
                        </div>
                      </div>
                    )}
                  </>
                ) : (
                  <div className="flex flex-col items-center justify-center py-20 text-center space-y-4">
                    <div className="w-20 h-20 rounded-full bg-emerald-50 flex items-center justify-center border-4 border-emerald-100"><CheckCircle2 className="w-10 h-10 text-emerald-500" /></div>
                    <div><h3 className="text-xl font-black text-slate-900 tracking-tight">SEM DIVERGÊNCIA</h3><p className="text-slate-500 font-medium">Todos os registros para esta métrica estão em conformidade.</p></div>
                  </div>
                )}
              </motion.div>
            )}
            {activeTab === 'objective' && (
              <motion.div key="objective" initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 10 }} transition={{ duration: 0.2 }} className="bg-slate-50 rounded-xl p-8 border border-slate-200">
                <div className="flex items-center gap-3 mb-6"><div className="p-3 bg-blue-100 text-blue-600 rounded-xl"><BookOpen className="w-8 h-8" /></div><h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Objetivo da Métrica</h3></div>
                <p className="text-lg text-slate-700 leading-relaxed font-medium">{metric.objective || "Nenhum objetivo detalhado cadastrado para esta métrica."}</p>
              </motion.div>
            )}
            {activeTab === 'rules' && (
              <motion.div key="rules" initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 10 }} transition={{ duration: 0.2 }} className="space-y-4">
                <div className="flex items-center gap-3 mb-4"><div className="p-3 bg-amber-100 text-amber-600 rounded-xl"><ShieldCheck className="w-8 h-8" /></div><h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Regras de Validação</h3></div>
                {metric.rules && metric.rules.length > 0 ? (
                  <div className="grid gap-3">{metric.rules.map((rule, idx) => (
                      <div key={idx} className="bg-white border border-slate-200 p-4 rounded-xl shadow-sm border-l-4 border-l-amber-500"><p className="text-slate-700 font-bold whitespace-pre-line text-sm leading-relaxed"><span className="text-amber-600 font-black mr-2">REGRA {idx + 1}:</span>{rule}</p></div>
                  ))}</div>
                ) : <p className="text-slate-500 italic p-8 bg-slate-50 rounded-xl border border-slate-200 text-center">Nenhuma regra de validação cadastrada para esta métrica.</p>}
              </motion.div>
            )}
            {activeTab === 'trend' && (
              <motion.div key="trend" initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 10 }} transition={{ duration: 0.2 }} className="space-y-6">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-3"><div className="p-3 bg-indigo-100 text-indigo-600 rounded-xl"><Activity className="w-8 h-8" /></div><h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Análise de Tendência</h3></div>
                  <div className="text-right"><p className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Valor Atual</p><p className="text-3xl font-black text-slate-900 tabular-nums">{metric.value}</p></div>
                </div>
                {metric.history && metric.history.length > 0 ? (
                  <div className="bg-white border border-slate-200 rounded-2xl p-6 shadow-sm h-[350px] relative">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={[...metric.history].reverse().map((val, idx) => ({ value: val, time: idx === metric.history!.length - 1 ? 'Agora' : `v-${metric.history!.length - idx - 1}` }))}>
                        <defs><linearGradient id="colorValue" x1="0" y1="0" x2="0" y2="1"><stop offset="5%" stopColor={metric.status === 'error' ? "#cc0000" : "#10b981"} stopOpacity={0.1}/><stop offset="95%" stopColor={metric.status === 'error' ? "#cc0000" : "#10b981"} stopOpacity={0}/></linearGradient></defs>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                        <XAxis dataKey="time" axisLine={false} tickLine={false} tick={{ fontSize: 10, fontWeight: 700, fill: '#94a3b8' }} dy={10} />
                        <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fontWeight: 700, fill: '#94a3b8' }} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontSize: '12px', fontWeight: 'bold' }} />
                        <Area type="monotone" dataKey="value" stroke={metric.status === 'error' ? "#cc0000" : "#10b981"} strokeWidth={4} fillOpacity={1} fill="url(#colorValue)" animationDuration={1500} />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                ) : <div className="bg-slate-50 border border-slate-200 rounded-xl p-12 text-center"><p className="text-slate-500 font-bold italic">Nenhum dado histórico disponível para esta métrica ainda.</p></div>}
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100"><p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Média</p><p className="text-lg font-black text-slate-800">{metric.history && metric.history.length > 0 ? (metric.history.reduce((acc, curr) => acc + curr, 0) / metric.history.length).toFixed(1) : '0'}</p></div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100"><p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Pico</p><p className="text-lg font-black text-slate-800">{metric.history && metric.history.length > 0 ? Math.max(...metric.history) : '0'}</p></div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100"><p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Status</p><p className={`text-lg font-black ${metric.status === 'error' ? 'text-brand-red' : 'text-emerald-500'}`}>{metric.status === 'error' ? 'CRÍTICO' : 'ESTÁVEL'}</p></div>
                  <div className="bg-slate-50 p-4 rounded-xl border border-slate-100"><p className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1">Meta</p><p className="text-lg font-black text-slate-800">0</p></div>
                </div>
              </motion.div>
            )}
            {activeTab === 'ai' && (
              <motion.div key="ai" initial={{ opacity: 0, x: -10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: 10 }} transition={{ duration: 0.2 }} className="space-y-6">
                <div className="flex items-center justify-between mb-4">
                  <div className="flex items-center gap-3"><div className="p-3 bg-indigo-100 text-indigo-600 rounded-xl"><Sparkles className="w-8 h-8" /></div><h3 className="text-2xl font-black text-slate-900 italic tracking-tighter uppercase">Inteligência Artificial</h3></div>
                  {aiAnalysis && <button onClick={runAiAnalysis} className="text-[10px] font-black text-indigo-600 uppercase tracking-widest flex items-center gap-1 hover:bg-indigo-50 px-2 py-1 rounded-lg transition-colors" disabled={isAnalyzing}><RefreshCcw className={`w-3 h-3 ${isAnalyzing ? 'animate-spin' : ''}`} />Refazer Análise</button>}
                </div>
                <div className="bg-slate-50 border border-slate-200 rounded-2xl p-8 min-h-[400px] relative overflow-y-auto max-h-[55vh]">
                  {isAnalyzing ? (
                    <div className="absolute inset-0 flex flex-col items-center justify-center gap-4 p-6 text-center">
                      <div className="relative"><motion.div animate={{ scale: [1, 1.2, 1], opacity: [0.3, 0.6, 0.3] }} transition={{ duration: 2, repeat: Infinity }} className="absolute inset-0 bg-indigo-200 rounded-full blur-xl"/><Sparkles className="w-12 h-12 text-indigo-600 relative animate-bounce" /></div>
                      <div><p className="font-black text-slate-800 uppercase tracking-widest text-sm">Gerando Insights com Gemini</p><p className="text-[10px] text-slate-500 font-bold uppercase mt-1 px-8">Isso pode levar alguns segundos enquanto analisamos padrões e regras...</p></div>
                    </div>
                  ) : aiAnalysis ? (
                    <div className="markdown-body prose prose-slate max-w-none prose-sm prose-headings:font-black prose-headings:uppercase prose-headings:italic prose-headings:tracking-tighter prose-p:text-slate-600 prose-p:font-medium prose-strong:text-slate-900"><Markdown>{aiAnalysis}</Markdown></div>
                  ) : (
                    <div className="flex flex-col items-center justify-center h-full text-center py-12 gap-6">
                      <div className="w-20 h-20 bg-indigo-100 rounded-full flex items-center justify-center shadow-inner"><Sparkles className="w-10 h-10 text-indigo-600" /></div>
                      <div className="max-w-md">
                        <h4 className="text-lg font-black text-slate-800 uppercase italic">Análise de Dados IA</h4>
                        <p className="text-slate-500 text-sm font-medium mt-2">Clique no botão forneceremos insights profundos sobre a tendência, causas raízes e recomendações para esta métrica.</p>
                        <button onClick={runAiAnalysis} className="mt-6 px-6 py-3 bg-indigo-600 text-white rounded-xl font-black uppercase tracking-widest text-sm shadow-lg shadow-indigo-200 hover:bg-indigo-700 transition-all active:scale-95 flex items-center gap-2 mx-auto"><Sparkles className="w-4 h-4" />Gerar Análise Agora</button>
                      </div>
                    </div>
                  )}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
        <footer className="px-6 py-4 bg-slate-50 border-t border-slate-100 flex justify-between items-center">
          <div className="flex items-center gap-2 text-slate-400"><Info className="w-4 h-4" /><span className="text-[10px] font-bold uppercase tracking-widest">Informações em tempo real do ERP/WMS</span></div>
          <button onClick={onClose} className="px-6 py-2 bg-slate-900 text-white rounded-lg font-black text-xs uppercase tracking-widest hover:bg-slate-800 transition-all active:scale-95 shadow-md">Fechar Painel</button>
        </footer>
      </motion.div>
    </motion.div>
  );
}

interface SupportModalProps {
  onClose: () => void;
  isWarRoom: boolean;
}

function SupportModal({ onClose, isWarRoom }: SupportModalProps) {
  const [copiedEmail, setCopiedEmail] = useState<'arlen' | 'digitalization' | null>(null);

  const handleCopy = (email: string, type: 'arlen' | 'digitalization') => {
    navigator.clipboard.writeText(email);
    setCopiedEmail(type);
    setTimeout(() => setCopiedEmail(null), 2000);
  };

  const handleEmailClick = (email: string, recipientName: string) => {
    const subject = `[Monitor Suporte] Contato / Dúvida`;
    const body = `Olá ${recipientName},

Gostaria de solicitar suporte referente ao Monitor do Dashboard.

--
Mensagem gerada via Central de Ajuda do Dashboard.`;

    window.location.href = `mailto:${email}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  };

  return (
    <motion.div 
      initial={{ opacity: 0 }} 
      animate={{ opacity: 1 }} 
      exit={{ opacity: 0 }} 
      className="fixed inset-0 z-[1100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
      onClick={onClose}
    >
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 15 }} 
        animate={{ opacity: 1, scale: 1, y: 0 }} 
        exit={{ opacity: 0, scale: 0.95, y: 15 }} 
        className={`rounded-2xl shadow-2xl w-full max-w-md overflow-hidden flex flex-col border transition-all ${
          isWarRoom ? 'bg-slate-900 border-slate-800 text-white' : 'bg-white border-slate-100 text-slate-900'
        }`}
        onClick={(e) => e.stopPropagation()}
      >
        <header className={`px-6 py-4 border-b flex justify-between items-center ${isWarRoom ? 'border-slate-800 bg-slate-950/40' : 'border-slate-100 bg-slate-50/50'}`}>
          <div className="flex items-center gap-3">
            <div className="w-1.5 h-7 rounded-full bg-indigo-600" />
            <div>
              <h2 className="text-base font-black uppercase tracking-tight italic flex items-center gap-1.5 leading-none">
                <HelpCircle className="w-5 h-5 text-indigo-600" /> Canais de Suporte
              </h2>
              <p className={`text-[10px] font-bold uppercase ${isWarRoom ? 'text-slate-400' : 'text-slate-500'}`}>Suporte & Canal DHL</p>
            </div>
          </div>
          <button onClick={onClose} className={`p-1.5 rounded-full transition-colors ${isWarRoom ? 'hover:bg-slate-800 text-slate-400' : 'hover:bg-slate-100 text-slate-500'}`}>
            <X className="w-5 h-5" />
          </button>
        </header>

        <div className="p-6 space-y-5">
          <p className={`text-xs leading-relaxed ${isWarRoom ? 'text-slate-400' : 'text-slate-500'}`}>
            Caso encontre alguma divergência ou precise de suporte técnico nas operações do dashboard, entre em contato direto com os responsáveis abaixo:
          </p>

          <div className="space-y-4">
            <div className={`p-4 rounded-xl border transition-all ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-200/50'}`}>
              <div className="flex justify-between items-start">
                <div className="cursor-pointer flex-1 mr-2" onClick={() => handleEmailClick('Arlen.Oliveira@dhl.com', 'Arlen Oliveira')}>
                  <h3 className="text-xs font-black uppercase text-indigo-600 tracking-wider hover:underline flex items-center gap-1.5">
                    Arlen Oliveira <Mail className="w-3.5 h-3.5" />
                  </h3>
                  <p className={`text-[10px] font-semibold mt-0.5 ${isWarRoom ? 'text-slate-400' : 'text-slate-500'}`}>Suporte Técnico & Operacional Specialist</p>
                  <p className="text-xs font-mono font-bold text-indigo-500 mt-2 select-all">Arlen.Oliveira@dhl.com</p>
                </div>
                <button 
                  type="button"
                  onClick={() => handleCopy('Arlen.Oliveira@dhl.com', 'arlen')}
                  className={`p-2 rounded-lg border transition-all ${
                    isWarRoom ? 'border-slate-800 bg-slate-900/60 hover:bg-slate-800' : 'border-slate-200 bg-white hover:bg-slate-50'
                  }`}
                  title="Copiar e-mail"
                >
                  {copiedEmail === 'arlen' ? <Check className="w-4 h-4 text-emerald-500" /> : <Copy className="w-4 h-4 text-slate-400" />}
                </button>
              </div>
            </div>

            <div className={`p-4 rounded-xl border transition-all ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-200/50'}`}>
              <div className="flex justify-between items-start">
                <div className="cursor-pointer flex-1 mr-2" onClick={() => handleEmailClick('innovationteam.processassistant@dhl.com', 'Time de Digitalização')}>
                  <h3 className="text-xs font-black uppercase text-indigo-600 tracking-wider hover:underline flex items-center gap-1.5">
                    Time de Digitalização <Mail className="w-3.5 h-3.5" />
                  </h3>
                  <p className={`text-[10px] font-semibold mt-0.5 ${isWarRoom ? 'text-slate-400' : 'text-slate-500'}`}>DHL Process Assistant Intelligence Team</p>
                  <p className="text-xs font-mono font-bold text-indigo-500 mt-2 select-all">innovationteam.processassistant@dhl.com</p>
                </div>
                <button 
                  type="button"
                  onClick={() => handleCopy('innovationteam.processassistant@dhl.com', 'digitalization')}
                  className={`p-2 rounded-lg border transition-all ${
                    isWarRoom ? 'border-slate-800 bg-slate-900/60 hover:bg-slate-800' : 'border-slate-200 bg-white hover:bg-slate-50'
                  }`}
                  title="Copiar e-mail"
                >
                  {copiedEmail === 'digitalization' ? <Check className="w-4 h-4 text-emerald-500" /> : <Copy className="w-4 h-4 text-slate-400" />}
                </button>
              </div>
            </div>
          </div>

          <div className={`p-4 text-center rounded-xl border ${isWarRoom ? 'border-slate-800 bg-slate-950/20 text-slate-500' : 'border-slate-100 bg-slate-50/50 text-slate-400'} text-[10px] italic leading-relaxed`}>
            Clique no nome do contato para abrir diretamente em seu Outlook ou copie o endereço usando os botões ao lado.
          </div>
        </div>
      </motion.div>
    </motion.div>
  );
}

export function Dashboard() {
  const [data, setData] = useState<Section[]>([]);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [selectedMetric, setSelectedMetric] = useState<Metric | null>(null);
  const [isWarRoom, setIsWarRoom] = useState(false);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isHelpOpen, setIsHelpOpen] = useState(false);
  const [settingsTab, setSettingsTab] = useState<'layout' | 'log' | 'prefs'>('layout');
  const [layoutConfig, setLayoutConfig] = useState<{title: string, width: number}[]>([]);
  const [eventLog, setEventLog] = useState<{ id: string, message: string, time: string, type: 'info' | 'critical' | 'success' }[]>([]);

  // System preferences (AI, header visibility, KPI details and audio notifications)
  const [preferences, setPreferences] = useState({
    enableAI: false,
    showHeader: true,
    showKpiDivergences: true,
    showKpiCritical: true,
    showKpiAccuracy: true,
    showKpiSla: true,
    audioAlertMode: 'none' as 'none' | 'beep' | 'ts' | 'both',
  });

  const [globalEmailAlertsEnabled, setGlobalEmailAlertsEnabled] = useState(false);
  const [globalTeamsAlertsEnabled, setGlobalTeamsAlertsEnabled] = useState(false);
  const [globalAiEnabled, setGlobalAiEnabled] = useState(false);

  // Helper to normalize strings for clean, readable URL parameters without accents or encoding issues
  const normalizeUrlParam = (text: string): string => {
    if (!text) return '';
    return text
      .toString()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9\s-]/g, '')
      .replace(/\s+/g, '-');
  };

  // States for operational filters with initial values parsed directly from URL parameters
  const [selectedDivision, setSelectedDivision] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      return params.get('divisao') || params.get('division') || 'TODAS';
    }
    return 'TODAS';
  });

  const [selectedSeverity, setSelectedSeverity] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      const sevParam = (params.get('severidade') || params.get('severity') || params.get('status') || 'TODOS').toLowerCase();
      if (['critical', 'critico', 'critica', 'error'].includes(sevParam)) return 'CRITICO';
      if (['stable', 'estavel', 'ok'].includes(sevParam)) return 'ESTAVEL';
    }
    return 'TODOS';
  });

  const [dashboardSearch, setDashboardSearch] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      return params.get('busca') || params.get('search') || '';
    }
    return '';
  });

  // Synchronize filter state back with Browser Address Bar URL parameters
  useEffect(() => {
    if (typeof window !== 'undefined') {
      const params = new URLSearchParams(window.location.search);
      
      if (selectedDivision && selectedDivision !== 'TODAS') {
        params.set('divisao', normalizeUrlParam(selectedDivision));
      } else {
        params.delete('divisao');
      }
      
      if (selectedSeverity && selectedSeverity !== 'TODOS') {
        params.set('severidade', selectedSeverity === 'CRITICO' ? 'critical' : 'stable');
      } else {
        params.delete('severidade');
      }
      
      if (dashboardSearch.trim()) {
        params.set('busca', dashboardSearch);
      } else {
        params.delete('busca');
      }
      
      const newRelativePathQuery = window.location.pathname + (params.toString() ? '?' + params.toString() : '');
      window.history.replaceState(null, '', newRelativePathQuery);
    }
  }, [selectedDivision, selectedSeverity, dashboardSearch]);

  // Handle browser Back / Forward navigation events (popstate)
  useEffect(() => {
    const handlePopState = () => {
      const params = new URLSearchParams(window.location.search);
      const divParam = params.get('divisao') || params.get('division');
      const sevParam = params.get('severidade') || params.get('severity') || params.get('status');
      const searchParam = params.get('busca') || params.get('search');
      
      if (divParam) {
        // Look up the actual case-sensitive/accented division name from data if loaded
        const normalizedDivParam = normalizeUrlParam(divParam);
        const matched = data.find(s => normalizeUrlParam(s.title) === normalizedDivParam);
        setSelectedDivision(matched ? matched.title : divParam);
      } else {
        setSelectedDivision('TODAS');
      }
      
      if (sevParam) {
        const normSev = sevParam.toLowerCase();
        if (['critical', 'critico', 'critica', 'error'].includes(normSev)) {
          setSelectedSeverity('CRITICO');
        } else if (['stable', 'estavel', 'ok'].includes(normSev)) {
          setSelectedSeverity('ESTAVEL');
        } else {
          setSelectedSeverity('TODOS');
        }
      } else {
        setSelectedSeverity('TODOS');
      }
      
      setDashboardSearch(searchParam || '');
    };
    
    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, [data]);

  // Update matched selectedDivision case-sensitive name of Section when Data loaded
  useEffect(() => {
    if (typeof window !== 'undefined' && data.length > 0) {
      const params = new URLSearchParams(window.location.search);
      const divParam = params.get('divisao') || params.get('division');
      if (divParam) {
        const normalizedDivParam = normalizeUrlParam(divParam);
        const found = data.find(s => normalizeUrlParam(s.title) === normalizedDivParam);
        if (found) {
          setSelectedDivision(found.title);
        } else {
          const foundPartial = data.find(s => s.title.toLowerCase().includes(divParam.toLowerCase()));
          if (foundPartial) {
            setSelectedDivision(foundPartial.title);
          }
        }
      }
    }
  }, [data]);

  useEffect(() => {
    const saved = localStorage.getItem('dashboard_preferences_v2');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Exclude legacy keys if they were stored in local storage
        delete parsed.enableEmailAlerts;
        delete parsed.enableTeamsAlerts;
        setPreferences(prev => ({ ...prev, ...parsed }));
      } catch (e) {
        console.error("Error loading dashboard preferences:", e);
      }
    }
  }, []);

  const savePreferences = (updated: typeof preferences) => {
    setPreferences(updated);
    localStorage.setItem('dashboard_preferences_v2', JSON.stringify(updated));
  };

  const triggerAlarm = (message: string, alertType: 'warning' | 'critical' | 'success') => {
    const mode = preferences.audioAlertMode;
    if (mode === 'none') return;
    
    if (mode === 'beep' || mode === 'both') {
      playAlertBeep(alertType);
    }
    if (mode === 'ts' || mode === 'both') {
      speakAlertText(message);
    }
  };

  const sendDivergenceEmail = async (metricTitle: string, details: any[]) => {
    if (!globalEmailAlertsEnabled) return;
    const currentCount = details ? details.length : 0;
    const lockKey = `${metricTitle}_${currentCount}`;

    // 1. Bloqueio Concorrente de Memória (E-mails disparados simultaneamente)
    if (sendingEmailLocks.has(lockKey)) {
      console.log(`[Email] Bloqueando envio duplicado concorrente em memória para "${metricTitle}" com ${currentCount} registros.`);
      return;
    }
    sendingEmailLocks.add(lockKey);

    const sentStorageKey = 'sent_divergence_alerts';
    let sentAlerts: Record<string, { lastSentAt: number; count: number }> = {};
    const nowTime = Date.now();

    try {
      console.log(`Iniciando envio de e-mail de alerta para a métrica: ${metricTitle}...`);
      
      // 2. Bloqueio Persistente (Evitar envio repetido nos últimos 30 minutos se a quantidade for idêntica)
      try {
        const stored = localStorage.getItem(sentStorageKey);
        if (stored) {
          sentAlerts = JSON.parse(stored);
        }
      } catch (e) {
        console.error("Erro ao ler alertas enviados do localStorage:", e);
      }

      const lastAlert = sentAlerts[metricTitle];
      if (lastAlert && lastAlert.count === currentCount && (nowTime - lastAlert.lastSentAt) < 30 * 60 * 1000) {
        console.log(`[Email] Ignorando envio para a métrica "${metricTitle}" - Alerta idêntico recente enviado há menos de 30 minutos.`);
        sendingEmailLocks.delete(lockKey);
        return;
      }

      // IMPORTANTE: Grava imediatamente no localStorage ANTES do processo assíncrono para blindar possíveis concorrências
      sentAlerts[metricTitle] = {
        lastSentAt: nowTime,
        count: currentCount
      };
      localStorage.setItem(sentStorageKey, JSON.stringify(sentAlerts));
      
      const alertRecipients = await fetchAlertEmails();
      if (!alertRecipients || alertRecipients.length === 0) {
        console.warn("Nenhum e-mail cadastrado na lista de alertas App_Dash_Emails. E-mail não será enviado.");
        sendingEmailLocks.delete(lockKey);
        return;
      }
      
      const emailListString = alertRecipients.map(u => u.email).filter(Boolean).join(';');
      if (!emailListString) {
        console.warn("Nenhum endereço de e-mail válido encontrado na lista de acessos.");
        sendingEmailLocks.delete(lockKey);
        return;
      }

      // Generate Excel attachment as Base64 using XLSX
      let attachmentsArray: { Name: string; ContentBytes: string }[] = [];
      if (details && details.length > 0) {
        try {
          const worksheet = XLSX.utils.json_to_sheet(details);
          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, "Divergências");
          const excelBase64 = XLSX.write(workbook, { bookType: 'xlsx', type: 'base64' });
          if (excelBase64) {
            attachmentsArray.push({
              Name: `Divergencias_${metricTitle.replace(/[^a-zA-Z0-9]/g, '_')}.xlsx`,
              ContentBytes: excelBase64
            });
            console.log("Anexo Excel gerado com sucesso em Base64!");
          }
        } catch (excelErr) {
          console.error("Erro ao gerar anexo Excel em Base64:", excelErr);
        }
      }

      // Determinar o Link completo do arquivo atual no SharePoint
      let currentSharepointPageUrl = window.location.origin;
      if (window._spPageContextInfo) {
        const ctx = window._spPageContextInfo;
        const webUrl = ctx.webAbsoluteUrl || ctx.siteAbsoluteUrl;
        const requestPath = ctx.serverRequestPath;
        if (webUrl && requestPath) {
          try {
            const urlObj = new URL(webUrl);
            currentSharepointPageUrl = urlObj.origin + requestPath;
          } catch {
            currentSharepointPageUrl = webUrl + requestPath;
          }
        } else if (webUrl) {
          currentSharepointPageUrl = webUrl;
        }
      } else {
        if (document.referrer && document.referrer.includes('.sharepoint.com')) {
          currentSharepointPageUrl = document.referrer;
        } else {
          currentSharepointPageUrl = window.location.href;
        }
      }

      // Build modern, stylish HTML body matching the requested layout design
      let tableRowsHtml = '';
      if (details && details.length > 0) {
        const columns = Object.keys(details[0]).slice(0, 5);
        
        const headerHtml = columns.map(c => `
          <th align="left" style="padding:12px 16px; border-bottom:2px solid #f3f4f6; font-size:10px; color:#6b7280; text-transform:uppercase; font-weight:bold; letter-spacing:1px;">${c}</th>
        `).join('');
        
        const rowsHtml = details.slice(0, 15).map((row, idx) => {
          const cells = columns.map(c => {
            const rawVal = row[c];
            const val = rawVal === null || rawVal === undefined ? '' : typeof rawVal === 'object' ? JSON.stringify(rawVal) : String(rawVal);
            return `<td style="padding:12px 16px; font-size:11px; color:#111827; line-height:1.5; border-bottom:1px solid #f3f4f6; font-family:monospace;">${val}</td>`;
          }).join('');
          const bg = idx % 2 === 0 ? '#ffffff' : '#f9fafb';
          return `<tr style="background-color: ${bg};">${cells}</tr>`;
        }).join('');

        const truncateWarning = details.length > 15 
          ? `<tr><td colspan="${columns.length}" style="text-align: center; padding: 14px 16px; color: #d40511; font-size: 11px; font-weight: bold; background-color: #fef2f2; border-top: 1px dashed #fee2e2;">Exibindo os 15 primeiros registros. A planilha completa (.xlsx) com todos os ${details.length} desvios foi anexada.</td></tr>`
          : `<tr><td colspan="${columns.length}" style="text-align: center; padding: 12px 16px; color: #10b981; font-size: 10px; font-weight: bold; background-color: #f0fdf4; border-top: 1px dashed #dcfce7; text-transform: uppercase;">Todos os ${details.length} registros listados acima. Planilha completa em anexo.</td></tr>`;

        tableRowsHtml = `
          <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse; border:1px solid #e5e7eb; border-radius:12px; overflow:hidden; background-color:#ffffff; margin-bottom:20px;">
            <thead>
              <tr style="background-color:#f9fafb;">
                ${headerHtml}
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
              ${truncateWarning}
            </tbody>
          </table>
        `;
      } else {
        tableRowsHtml = `<p style="color: #64748b; font-style: italic; background-color: #f8fafc; padding: 16px; border-radius: 8px; border: 1px dashed #cbd5e1; font-size: 13px; text-align: center; margin-bottom:20px;">Nenhum detalhe adicional disponível para esta divergência.</p>`;
      }

      const bodyHtml = `
        <!DOCTYPE html>
        <html lang="pt-BR">
        <head>
          <meta charset="UTF-8">
          <title>Divergência Detectada</title>
        </head>
        <body style="margin:0; padding:0; background-color:#f8f9fa; font-family:Arial, Helvetica, sans-serif;">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f8f9fa; border-collapse:collapse;">
            <tr>
              <td align="center" style="padding:40px 10px;">
                <table border="0" cellpadding="0" cellspacing="0" width="600" style="background-color:#ffffff; border-collapse:collapse; border-radius:24px; border:1px solid #e5e7eb; overflow:hidden; box-shadow:0 4px 6px -1px rgba(0,0,0,0.05);">

                  <tr>
                    <td style="background-color:#1a1a1a; padding:0;">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                        <tr>
                          <td height="6" bgcolor="#d40511" style="font-size:1px; line-height:6px;">&nbsp;</td>
                        </tr>
                        <tr>
                          <td style="padding:30px 40px;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                              <tr>
                                <td>
                                  <div style="color:#ffcc00; font-size:10px; font-weight:900; text-transform:uppercase; letter-spacing:2px; margin-bottom:8px;">ALERTA DE SISTEMA</div>
                                  <h1 style="color:#ffffff; margin:0; font-size:24px; font-weight:bold; text-transform:uppercase; line-height:1.2; letter-spacing:-1px;">
                                    Divergência <br><span style="color:#d40511;">Detectada</span>
                                  </h1>
                                </td>
                                <td align="right" style="vertical-align:middle;">
                                  <div style="border:1px solid #333333; padding:8px 12px; border-radius:8px; background-color:#262626; color:#ffffff; font-size:10px; font-weight:bold; text-transform:uppercase; letter-spacing:1px;">
                                    SLAs Monitor
                                  </div>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>

                  <tr>
                    <td style="padding:40px;">
                      <p style="margin:0 0 16px 0; font-size:18px; color:#111827; font-weight:bold;">
                        Olá Equipe,
                      </p>

                      <p style="margin:0 0 32px 0; font-size:15px; color:#4b5563; line-height:1.6;">
                        O sistema identificou divergências no monitoramento operacional da métrica selecionada. Por favor, revise os dados abaixo e as planilhas completas enviadas em anexo.
                      </p>

                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#ffffff; border-collapse:collapse; border:1px solid #f3f4f6; border-radius:16px; margin-bottom:32px;">
                        <tr>
                          <td style="padding:24px; border-radius:16px; border:1px solid #f3f4f6;">
                            <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                              <tr>
                                <td style="padding-bottom:20px;">
                                  <div style="font-size:10px; color:#9ca3af; text-transform:uppercase; font-weight:900; letter-spacing:1px; margin-bottom:6px;">
                                    MÉTRICA MONITORADA
                                  </div>
                                  <div style="font-size:16px; color:#111827; font-weight:bold; word-break:break-all;">
                                    ${metricTitle}
                                  </div>
                                </td>
                              </tr>
                              <tr>
                                <td>
                                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                                    <tr>
                                      <td width="100%" style="vertical-align:top;">
                                        <div style="font-size:10px; color:#9ca3af; text-transform:uppercase; font-weight:900; letter-spacing:1px; margin-bottom:6px;">
                                          QUANTIDADE DE DIVERGÊNCIAS
                                        </div>
                                        <div style="font-size:18px; color:#d40511; font-weight:bold;">
                                          ${currentCount} registros de desvio
                                        </div>
                                      </td>
                                    </tr>
                                  </table>
                                </td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>

                      <div style="margin-bottom:12px;">
                        <span style="font-size:12px; color:#111827; text-transform:uppercase; font-weight:900; letter-spacing:1px;">
                          Amostra das Divergências (MÁXIMO 15)
                        </span>
                      </div>

                      ${tableRowsHtml}

                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color:#f9fafb; border-collapse:collapse; border-radius:12px; border:1px solid #f3f4f6; margin-top:32px;">
                        <tr>
                          <td style="padding:24px;">
                            <p style="margin:0; font-size:13px; color:#6b7280; line-height:1.6;">
                              <strong style="color:#111827;">Ação necessária:</strong>
                              A planilha Excel (.xlsx) com a listagem completa com todos os desvios foi gerada pelo sistema e anexada a este e-mail para análise integrada. Você também pode visualizar os dados atualizados clicando no link abaixo.
                            </p>
                            <div style="margin-top:20px; text-align:center;">
                              <a href="${currentSharepointPageUrl}" style="display:inline-block; background-color:#d40511; color:#ffffff; font-size:12px; font-weight:bold; padding:12px 24px; text-decoration:none; border-radius:8px; text-transform:uppercase; letter-spacing:1px; box-shadow:0 4px 10px rgba(212,5,17,0.25);">
                                Acessar Painel Online
                              </a>
                            </div>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>

                  <tr>
                    <td style="padding:0 40px 40px 40px;">
                      <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-top:1px solid #f3f4f6; padding-top:32px; border-collapse:collapse;">
                        <tr>
                          <td align="center">
                            <p style="margin:0; font-size:11px; color:#9ca3af; text-transform:uppercase; font-weight:900; letter-spacing:2px;">
                              DHL Supply Chain • SISTEMA MONITOR OPERACIONAL
                            </p>
                            <p style="margin:4px 0 0 0; font-size:10px; color:#94a3b8;">
                              Este e-mail é gerado automaticamente pelo robô de monitoramento de SLAs. Não responda a esta mensagem.
                            </p>
                            <p style="margin:8px 0 0 0; font-size:10px; color:#94a3b8; font-family:monospace;">
                              Data e hora da detecção: ${new Date().toLocaleString('pt-BR')} (UTC)
                            </p>
                          </td>
                        </tr>
                      </table>
                    </td>
                  </tr>

                </table>
              </td>
            </tr>
          </table>
        </body>
        </html>
      `;

      const Title = `[Alerta Operacional] Divergência em: ${metricTitle}`;

      const directEmailUrl = process.env.POWER_AUTOMATE_EMAIL_URL;
      const emailEndpoint = directEmailUrl && directEmailUrl !== "MY_POWER_AUTOMATE_EMAIL_URL"
        ? directEmailUrl
        : "/api/send-email";

      console.log(`Enviando e-mail de divergência via: ${emailEndpoint === directEmailUrl ? 'Link Direto (Vite Exposta)' : 'Proxy Local Express'}`);

      const response = await fetch(emailEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          emails: emailListString,
          Title,
          BodyEmail: bodyHtml,
          Attachments: attachmentsArray
        })
      });

      if (!response.ok) {
        throw new Error(`Servidor de e-mail retornou status: ${response.status}`);
      }

      console.log(`E-mail com planilha Excel anexada enviado com sucesso para: ${emailListString}`);
      setEventLog(prev => ([{ id: Math.random().toString(36).substr(2, 9), message: `E-mail de desvio enviado à lista com planilha Excel anexada para "${metricTitle}".`, time: new Date().toLocaleTimeString('pt-BR'), type: 'success' as const }, ...prev] as any).slice(0, 50));
    } catch (error: any) {
      console.error("Erro ao tentar enviar e-mail via Power Automate:", error);
      // Remove do localStorage para permitir nova tentativa em caso de erro absoluto de envio
      try {
        const stored = localStorage.getItem(sentStorageKey);
        if (stored) {
          const sent = JSON.parse(stored);
          delete sent[metricTitle];
          localStorage.setItem(sentStorageKey, JSON.stringify(sent));
        }
      } catch (cleanErr) {}
      setEventLog(prev => ([{ id: Math.random().toString(36).substr(2, 9), message: `Falha no envio de e-mail (${metricTitle}): ${error.message || error}`, time: new Date().toLocaleTimeString('pt-BR'), type: 'critical' as const }, ...prev] as any).slice(0, 50));
    } finally {
      // Manter o lock temporário em memória por 10 segundos adicionais para amparar flutuações e sincronizações em série
      setTimeout(() => {
        sendingEmailLocks.delete(lockKey);
      }, 10000);
    }
  };

  const sendingTeamsLocks = new Set<string>();

  const sendTeamsNotification = async (metricTitle: string, details: any[]) => {
    if (!globalTeamsAlertsEnabled) return;
    const currentCount = details ? details.length : 0;
    const lockKey = `${metricTitle}_${currentCount}`;

    // 1. Bloqueio Concorrente de Memória para evitar envios duplicados repetidos
    if (sendingTeamsLocks.has(lockKey)) {
      console.log(`[Teams] Bloqueando envio duplicado concorrente para "${metricTitle}" com ${currentCount} registros.`);
      return;
    }
    sendingTeamsLocks.add(lockKey);

    const sentTeamsStorageKey = 'sent_teams_alerts';
    try {
      // 2. Bloqueio por Sincronização de Data no localStorage (Evitar re-envios frequentes no mesmo período)
      const stored = localStorage.getItem(sentTeamsStorageKey);
      let sentAlerts: Record<string, string> = {};
      if (stored) {
        try {
          sentAlerts = JSON.parse(stored);
        } catch (e) {
          sentAlerts = {};
        }
      }

      // Se já disparou este card há menos de 10 minutos com o mesmo número de erros, ignora
      const lastSentTime = sentAlerts[metricTitle];
      if (lastSentTime) {
        const diffMs = Date.now() - new Date(lastSentTime).getTime();
        const tenMinutes = 10 * 60 * 1000;
        if (diffMs < tenMinutes) {
          console.log(`[Teams] Alerta do Teams para "${metricTitle}" enviado recentemente (${Math.round(diffMs / 1000)}s atrás). Ignorando para evitar flood.`);
          sendingTeamsLocks.add(lockKey); // Manter bloqueado
          return;
        }
      }

      // Registra no cache que está enviando agora
      sentAlerts[metricTitle] = new Date().toISOString();
      localStorage.setItem(sentTeamsStorageKey, JSON.stringify(sentAlerts));

      console.log(`[Teams] Iniciando envio de desvio via Teams para o card "${metricTitle}"...`);

      const directTeamsUrl = process.env.POWER_AUTOMATE_TEAMS_URL;
      const teamsEndpoint = directTeamsUrl && directTeamsUrl !== "MY_POWER_AUTOMATE_TEAMS_URL"
        ? directTeamsUrl
        : "https://51a805d34213e248a3506f5db8fe28.55.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/d6034ec10bbe4c39b8e9b2e3d03f5e5a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=miH_7ZEESA8UAeFa3_KUvuiij_05CJxaI0zwVIeufqA";

      const teamsChatId = await getTeamsChatId();

      // 3. Gerar print de altíssima qualidade do painel usando html2canvas em um clone isolado para privacidade do usuário e foco operacional
      let base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="; //Fallback de 1x1 png transparente
      const targetElement = document.getElementById('main-dashboard') || document.body;
      
      if (targetElement) {
        let safeCloneResult: { clone: HTMLElement; cleanup: () => void } | null = null;
        try {
          // Criar clone do Dashboard fora da tela (em left: -15000px) para total isolamento
          safeCloneResult = createHtml2CanvasSafeClone(targetElement, isWarRoom);
          const cloneWrapper = safeCloneResult.clone;
          const clonedDashboard = cloneWrapper.querySelector('#main-dashboard') || cloneWrapper.firstChild as HTMLElement;

          if (clonedDashboard) {
            // Encontrar qual divisão (section) no data possui a métrica que está sendo reportada
            const targetSec = data.find(s => 
              s.metrics.some(m => m.title.trim().toLowerCase() === metricTitle.trim().toLowerCase())
            );
            
            // Buscar todas as seções (division containers) dentro do clone
            const clonedSecs = Array.from(clonedDashboard.querySelectorAll('section')) as HTMLElement[];
            
            if (targetSec) {
              clonedSecs.forEach(clonedSec => {
                const h2El = clonedSec.querySelector('h2');
                const clonedSecTitle = h2El ? h2El.textContent?.trim().toLowerCase() : '';
                const targetTitleLower = targetSec ? targetSec.title.trim().toLowerCase() : '';
                
                if (clonedSecTitle !== targetTitleLower) {
                  // Remover divisões que não sejam a divisão alvo do alerta
                  clonedSec.parentNode?.removeChild(clonedSec);
                } else {
                  // Na nossa divisão alvo, manter o card reportado (mesmo se estiver OK) e todos os críticos, remover os estáveis
                  const cardElements = Array.from(clonedSec.querySelectorAll('[id^="card-"]')) as HTMLElement[];
                  cardElements.forEach(cardEl => {
                    const cardId = cardEl.id;
                    const metricId = cardId.replace('card-', '');
                    const metricObj = targetSec?.metrics.find(m => String(m.id) === metricId);
                    
                    if (metricObj) {
                      const isReportedCard = metricObj.title.trim().toLowerCase() === metricTitle.trim().toLowerCase();
                      const isCritical = metricObj.status === 'error';
                      if (!isReportedCard && !isCritical) {
                        cardEl.parentNode?.removeChild(cardEl);
                      }
                    } else {
                      // Fallback por comparação de título caso ID mude
                      const h3El = cardEl.querySelector('h3');
                      const cardTitle = h3El ? h3El.textContent?.trim().toLowerCase() : '';
                      const targetTitleLower = metricTitle.trim().toLowerCase();
                      const isReportedCard = cardTitle === targetTitleLower;
                      
                      const metricObjByTitle = targetSec?.metrics.find(m => m.title.trim().toLowerCase() === cardTitle);
                      const isCritical = metricObjByTitle ? metricObjByTitle.status === 'error' : false;
                      
                      if (!isReportedCard && !isCritical) {
                        cardEl.parentNode?.removeChild(cardEl);
                      }
                    }
                  });

                  // Ocultar botões de navegação lateral ou de rolagem deste container clonado
                  const scrollIndicator = clonedSec.querySelector('.flex.items-center.justify-between.mt-2.px-4.w-full');
                  if (scrollIndicator) {
                    scrollIndicator.parentNode?.removeChild(scrollIndicator);
                  }
                }
              });
            } else {
              // Fallback: Se não encontrarmos a divisão pelo título, filtramos todas as divisões para mostrar apenas o card reportado e os críticos
              clonedSecs.forEach(clonedSec => {
                const h2El = clonedSec.querySelector('h2');
                const clonedSecTitle = h2El ? h2El.textContent?.trim().toLowerCase() : '';
                const matchedSecData = data.find(s => s.title.trim().toLowerCase() === clonedSecTitle);

                const cardElements = Array.from(clonedSec.querySelectorAll('[id^="card-"]')) as HTMLElement[];
                let visibleCount = 0;
                cardElements.forEach(cardEl => {
                  const cardId = cardEl.id;
                  const metricId = cardId.replace('card-', '');
                  const metricObj = matchedSecData?.metrics.find(m => String(m.id) === metricId);
                  
                  if (metricObj) {
                    const isReportedCard = metricObj.title.trim().toLowerCase() === metricTitle.trim().toLowerCase();
                    const isCritical = metricObj.status === 'error';
                    if (isReportedCard || isCritical) {
                      visibleCount++;
                    } else {
                      cardEl.parentNode?.removeChild(cardEl);
                    }
                  } else {
                    const h3El = cardEl.querySelector('h3');
                    const cardTitle = h3El ? h3El.textContent?.trim().toLowerCase() : '';
                    const isReportedCard = cardTitle === metricTitle.trim().toLowerCase();
                    if (isReportedCard) {
                      visibleCount++;
                    } else {
                      cardEl.parentNode?.removeChild(cardEl);
                    }
                  }
                });

                if (visibleCount === 0) {
                  clonedSec.parentNode?.removeChild(clonedSec);
                } else {
                  const scrollIndicator = clonedSec.querySelector('.flex.items-center.justify-between.mt-2.px-4.w-full');
                  if (scrollIndicator) {
                    scrollIndicator.parentNode?.removeChild(scrollIndicator);
                  }
                }
              });
            }
            
            // Remover painéis de controle, botões de filtros globais etc. no clone
            const anyFilters = clonedDashboard.querySelectorAll('.mx-auto.w-full.p-4.rounded-2xl.border');
            anyFilters.forEach(f => f.parentNode?.removeChild(f));
          }

          const scaleStr = process.env.TEAMS_SCREENSHOT_SCALE;
          const captureScale = scaleStr ? parseFloat(scaleStr) : 2.5;

          // Tirar o print do wrapper fora de tela (cloneWrapper) totalmente limpo e isolado
          const canvas = await withCleanedEnvironment(async () => {
            return await html2canvas(cloneWrapper, {
              useCORS: true,
              allowTaint: true,
              backgroundColor: isWarRoom ? '#050510' : '#f8fafc',
              scale: captureScale, 
              logging: false,
              imageTimeout: 0,
              scrollX: 0,
              scrollY: 0,
            });
          });

          const imgData = canvas.toDataURL('image/png');
          if (imgData && imgData.includes('base64,')) {
            base64Image = imgData.split('base64,')[1];
          }
        } catch (screenshotErr) {
          console.error("[Teams] Erro ao obter print isolado da tela via html2canvas:", screenshotErr);
        } finally {
          // Destruir e limpar o clone do body imediatamente
          if (safeCloneResult) {
            safeCloneResult.cleanup();
          }
        }
      }

      // 4. Estruturar tabela detalhada para as primeiras linhas
      const rowsHtml = details && details.length > 0 
        ? details.slice(0, 8).map((row, rIdx) => {
            const keys = Object.keys(row).slice(0, 5); // limitar a 5 colunas principais
            return `<tr style="border-bottom: 1px solid #e2e8f0; ${rIdx % 2 === 0 ? 'background-color: #f8fafc;' : ''}">
              ${keys.map(k => `<td style="padding: 8px 10px; font-size: 11px; color: #334155; font-family: 'Segoe UI', Arial;">${String(row[k] ?? '')}</td>`).join('')}
            </tr>`;
          }).join('')
        : '';

      const headersHtml = details && details.length > 0
        ? `<tr style="background-color: #f1f5f9; font-weight: bold; text-align: left; border-bottom: 2px solid #cbd5e1;">
            ${Object.keys(details[0]).slice(0, 5).map(k => `<th style="padding: 8px 10px; font-size: 11px; text-transform: uppercase; color: #475569; font-family: 'Segoe UI', Arial;">${k}</th>`).join('')}
           </tr>`
        : '';

      const detailsTable = rowsHtml 
        ? `<table style="width: 100%; border-collapse: collapse; margin-top: 12px; margin-bottom: 18px; border: 1px solid #e2e8f0;">
            <thead>${headersHtml}</thead>
            <tbody>${rowsHtml}</tbody>
           </table>`
        : '<p style="color: #64748b; font-size: 12px; font-family: \'Segoe UI\', Arial;">Nenhuma linha de detalhe disponível.</p>';

      const htmlBody = `
        <div style="font-family: 'Segoe UI', Arial, sans-serif; max-width: 650px; border: 1px solid #cbd5e1; border-radius: 12px; overflow: hidden; background-color: #ffffff;">
          <div style="background-color: #8b0000; color: #ffffff; padding: 18px 24px;">
            <h2 style="margin: 0; font-size: 18px; text-transform: uppercase; font-style: italic; font-weight: 900; letter-spacing: -0.5px; font-family: 'Segoe UI', Arial;">🚨 Alerta de Desvio Operacional</h2>
            <p style="margin: 4px 0 0 0; font-size: 11px; opacity: 0.85; font-weight: bold; letter-spacing: 0.5px;">Painel de Controle em Tempo Real</p>
          </div>
          <div style="padding: 24px;">
            <p style="font-size: 14px; margin-top: 0; color: #1e293b; font-family: 'Segoe UI', Arial;">Olá, equipe!</p>
            <p style="font-size: 14px; color: #334155; font-family: 'Segoe UI', Arial; line-height: 1.5;">
              Identificamos um comportamento impeditivo no card de monitoramento <strong>"${metricTitle}"</strong>. Seguem abaixo as informações da ocorrência e a captura em tempo real.
            </p>
            
            <div style="background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 12px 18px; margin: 20px 0; border-radius: 6px;">
              <p style="margin: 0; font-size: 13px; color: #991b1b; font-weight: 900; font-family: 'Segoe UI', Arial;">
                Métrica: ${metricTitle}
              </p>
              <p style="margin: 4px 0 0 0; font-size: 12px; color: #7f1d1d; font-family: 'Segoe UI', Arial; font-weight: 500;">
                Divergências ativas: <strong>${currentCount}</strong> registros.
              </p>
            </div>

            <h3 style="font-size: 12px; font-weight: 900; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 24px; margin-bottom: 6px; color: #1e293b; font-family: 'Segoe UI', Arial;">Resumo dos Dados</h3>
            ${detailsTable}

            <p style="font-size: 10px; color: #94a3b8; margin-top: 30px; border-top: 1px solid #f1f5f9; padding-top: 15px; font-weight: bold; font-family: 'Segoe UI', Arial;">
              ESTA É UMA NOTIFICAÇÃO AUTOMÁTICA DISPARADA PARA O MS TEAMS (CHAT ID: ${teamsChatId}). POR FAVOR, NÃO RESPONDA DIRETAMENTE.
            </p>
          </div>
        </div>
      `;

      const teamsPayload = {
        destinoTipo: "grupo",
        chatId: teamsChatId,
        html: htmlBody,
        imagens: [
          {
            nome: "print-dashboard.png",
            base64: base64Image,
            contentType: "image/png",
            alt: "Print do Dashboard em Tempo Real"
          }
        ]
      };

      const response = await fetch(teamsEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(teamsPayload)
      });

      if (!response.ok) {
        throw new Error(`Teams API respondeu com status: ${response.status}`);
      }

      console.log(`[Teams] Alerta Teams enviado com sucesso para "${metricTitle}" no chat ${teamsChatId}`);
      setEventLog(prev => ([{ id: Math.random().toString(36).substr(2, 9), message: `Alerta Teams de divergência enviado para "${metricTitle}".`, time: new Date().toLocaleTimeString('pt-BR'), type: 'success' as const }, ...prev] as any).slice(0, 50));
    } catch (error: any) {
      console.error("[Teams] Erro ao tentar enviar mensagem para o Teams:", error);
      // Limpa do cache para nova tentativa
      try {
        const stored = localStorage.getItem(sentTeamsStorageKey);
        if (stored) {
          const sent = JSON.parse(stored);
          delete sent[metricTitle];
          localStorage.setItem(sentTeamsStorageKey, JSON.stringify(sent));
        }
      } catch (cleanErr) {}
      setEventLog(prev => ([{ id: Math.random().toString(36).substr(2, 9), message: `Falha no alerta Teams (${metricTitle}): ${error.message || error}`, time: new Date().toLocaleTimeString('pt-BR'), type: 'critical' as const }, ...prev] as any).slice(0, 50));
    } finally {
      // Libera o lock de concorrência breve em memória
      setTimeout(() => {
        sendingTeamsLocks.delete(lockKey);
      }, 5000);
    }
  };

  // Permissions & Access Management States
  const [hasAdminAccess, setHasAdminAccess] = useState(false);
  const [isAccessModalOpen, setIsAccessModalOpen] = useState(false);
  const [allowedUsers, setAllowedUsers] = useState<{ id: string; email: string }[]>([]);
  const [newAccessEmail, setNewAccessEmail] = useState('');
  const [isSavingAccess, setIsSavingAccess] = useState(false);
  const [accessMessage, setAccessMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

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
        
        const email = getCurrentSharePointUserEmail() || localStorage.getItem('mock_user_email') || 'arlenloran@gmail.com';
        const allowed = await isUserAllowed(email);
        setHasAdminAccess(allowed);
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
          const allowed = await isUserAllowed(curEmail);
          setHasAdminAccess(allowed);
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

  useEffect(() => {
    if (layoutConfig.length > 0) {
      localStorage.setItem('dashboard_layout', JSON.stringify(layoutConfig));
    }
  }, [layoutConfig]);

  useEffect(() => {
    const initApp = async () => {
      await ensureSharePointConfig();
      const config = await fetchDashboardConfig();

      // Detect expired metrics on mount to display them instantly in a stable green state, then load background queries concurrently
      const expiredList: string[] = [];
      const expiredConfig = config.map(section => {
        const updatedMetrics = section.metrics.map(metric => {
          let isExpired = false;
          if (!metric.lastUpdateAt) {
            isExpired = true;
          } else {
            const last = new Date(metric.lastUpdateAt);
            if (isNaN(last.getTime())) {
              isExpired = true;
            } else {
              const expiryTime = last.getTime() + (metric.refreshInterval || 5) * 60000;
              isExpired = Date.now() >= expiryTime;
            }
          }

          if (isExpired && metric.sqlQuery) {
            expiredList.push(metric.id);
            // Return temporary clean, stable green state (no discrepancies) so card appears on screen immediately as green/0
            return {
              ...metric,
              status: 'ok' as const,
              value: 0,
              lastUpdate: 'Sincronizando...',
              lastUpdateAt: new Date().toISOString() // temporary to prevent standard Countdown from triggering individual immediate refreshes on mount
            };
          }
          return metric;
        });
        return { ...section, metrics: updatedMetrics };
      });

      setData(expiredConfig);

      // Verify if current user is admin
      const email = getCurrentSharePointUserEmail() || localStorage.getItem('mock_user_email') || 'arlenloran@gmail.com';
      const allowed = await isUserAllowed(email);
      setHasAdminAccess(allowed);
      
      const savedLayout = localStorage.getItem('dashboard_layout');
      let currentLayout: {title: string, width: number}[] = [];
      
      if (savedLayout) {
        try {
          const parsed = JSON.parse(savedLayout);
          if (Array.isArray(parsed)) {
            // Filter out sections that are no longer in the retrieved list from database/SharePoint config
            currentLayout = parsed.filter(c => config.some(s => s.title === c.title));
          }
          // Add missing sections from config to layout
          config.forEach(s => {
            if (!currentLayout.find(c => c.title === s.title)) {
              currentLayout.push({ title: s.title, width: 33.33 });
            }
          });
        } catch (e) {
          currentLayout = config.map(s => ({ title: s.title, width: 33.33 }));
        }
      } else {
        currentLayout = config.map(s => ({ title: s.title, width: 33.33 }));
      }
      
      setLayoutConfig(currentLayout);

      try {
        const emailE = await getEmailAlertsEnabled();
        setGlobalEmailAlertsEnabled(emailE);
      } catch (err) {
        console.error("Error loading EmailAlertsEnabled in init:", err);
      }

      try {
        const teamsE = await getTeamsAlertsEnabled();
        setGlobalTeamsAlertsEnabled(teamsE);
      } catch (err) {
        console.error("Error loading TeamsAlertsEnabled in init:", err);
      }

      try {
        const aiE = await getAiEnabled();
        setGlobalAiEnabled(aiE);
      } catch (err) {
        console.error("Error loading AiEnabled in init:", err);
      }

      // Concurrently query only the expired cards in background, updating state progressively as they arrive!
      if (expiredList.length > 0) {
        console.log(`Starting background progressive updates for ${expiredList.length} expired metrics all at once...`);
        
        // Find them in original config to get proper initial state context
        const expiredMetrics = config.flatMap(sec => 
          sec.metrics.filter(m => expiredList.includes(m.id) && m.sqlQuery)
        );

        expiredMetrics.forEach((metric, index) => {
          postSqlQuery(metric.sqlQuery!, metric.id)
            .then(result => {
              const resolveNow = new Date();
              const resolveNowIso = resolveNow.toISOString();
              const rowCount = Array.isArray(result) ? result.length : 0;
              const isNowError = rowCount > 0;
              const wasError = metric.status === 'error'; // original database state status

              const updatedHistory = [rowCount, ...(metric.history || [])].slice(0, 10);

              // Update metrics array in data state progressively as results arrive
              setData(current => {
                return current.map(section => {
                  const mIdx = section.metrics.findIndex(m => m.id === metric.id);
                  if (mIdx !== -1) {
                    const currentMetric = section.metrics[mIdx];
                    const updatedMetrics = [...section.metrics];
                    updatedMetrics[mIdx] = {
                      ...currentMetric,
                      value: rowCount,
                      details: Array.isArray(result) ? result : [],
                      lastUpdate: resolveNow.toLocaleString('pt-BR'),
                      lastUpdateAt: resolveNowIso, // next countdown timer set properly from execution timestamp
                      status: isNowError ? 'error' : 'ok',
                      history: updatedHistory
                    };
                    return { ...section, metrics: updatedMetrics };
                  }
                  return section;
                });
              });

              // Save to SharePoint asynchronously. Stagger calls to prevent concurrent lock clashes!
              setTimeout(() => {
                saveMetricData(metric.id, resolveNowIso, result, updatedHistory)
                  .catch(err => console.error(`Error saving background metric data for ${metric.id}:`, err));
              }, index * 1000);

              // Add success entry into administrative logs list
              setEventLog(prev => ([{
                id: Math.random().toString(36).substr(2, 9),
                message: `Métrica "${metric.title}" sincronizada com sucesso.`,
                time: resolveNow.toLocaleTimeString('pt-BR'),
                type: 'success' as const
              }, ...prev] as any).slice(0, 50));

              // State transitions alert routing rules
              if (isNowError && !wasError) {
                triggerAlarm(`Alerta! Nova divergência detectada na métrica: ${metric.title}`, 'critical');
                try {
                  sendDivergenceEmail(metric.title, Array.isArray(result) ? result : []);
                } catch (e1) {
                  console.error("Error sending email:", e1);
                }
                try {
                  sendTeamsNotification(metric.title, Array.isArray(result) ? result : []);
                } catch (e2) {
                  console.error("Error sending Teams notification:", e2);
                }
              } else if (!isNowError && wasError) {
                triggerAlarm(`Excelente! Divergência resolvida na métrica: ${metric.title}`, 'success');
              }
            })
            .catch(err => {
              console.error(`Error refreshing expired card background query for ${metric.title}:`, err);
              const resolveNow = new Date();
              // In case query fails, mark as error synchronizing so user registers discrepancy
              setData(current => {
                return current.map(section => {
                  const mIdx = section.metrics.findIndex(m => m.id === metric.id);
                  if (mIdx !== -1) {
                    const currentMetric = section.metrics[mIdx];
                    const updatedMetrics = [...section.metrics];
                    updatedMetrics[mIdx] = {
                      ...currentMetric,
                      lastUpdate: `Erro de sincronização (${resolveNow.toLocaleTimeString('pt-BR')})`,
                      lastUpdateAt: resolveNow.toISOString() // reset timer trigger
                    };
                    return { ...section, metrics: updatedMetrics };
                  }
                  return section;
                });
              });
            });
        });
      }
    };
    initApp();
  }, []);

  const refreshData = async () => {
    setIsRefreshing(true);
    try {
      const config = await fetchDashboardConfig();
      setData(config);
      
      // Update layout config if new sections appeared or old ones were removed
      setLayoutConfig(prev => {
        // Keep only those that are currently present in the active configuration
        const activeConfigs = prev.filter(c => config.some(s => s.title === c.title));
        config.forEach(s => {
          if (!activeConfigs.find(c => c.title === s.title)) {
            activeConfigs.push({ title: s.title, width: 33.33 });
          }
        });
        return activeConfigs;
      });

      try {
        const emailE = await getEmailAlertsEnabled();
        setGlobalEmailAlertsEnabled(emailE);
      } catch (err) {
        console.error("Error loading EmailAlertsEnabled in refresh:", err);
      }

      try {
        const teamsE = await getTeamsAlertsEnabled();
        setGlobalTeamsAlertsEnabled(teamsE);
      } catch (err) {
        console.error("Error loading TeamsAlertsEnabled in refresh:", err);
      }

      try {
        const aiE = await getAiEnabled();
        setGlobalAiEnabled(aiE);
      } catch (err) {
        console.error("Error loading AiEnabled in refresh:", err);
      }

      await fetchAllDataInternal(config);
    } catch (err) {
      console.error("Refresh error:", err);
    } finally {
      setIsRefreshing(false);
    }
  };

  const fetchAllDataInternal = async (configRef: Section[]) => {
    if (configRef.length === 0) return;

    const metricsToUpdate: { sectionTitle: string, metric: Metric }[] = [];
    configRef.forEach(section => {
      section.metrics.forEach(metric => {
        if (metric.sqlQuery) {
          metricsToUpdate.push({ sectionTitle: section.title, metric });
        }
      });
    });

    if (metricsToUpdate.length === 0) return;

    // Put metrics in a visible visual synchronization state instantly, 
    // ensuring those that are expired render stable and green initially to remove any load lag!
    setData(current => {
      return current.map(section => {
        const updatedMetrics = section.metrics.map(metric => {
          const matched = metricsToUpdate.some(m => m.metric.id === metric.id);
          if (matched) {
            // Check if metric is expired relative to now
            let isExpired = false;
            if (!metric.lastUpdateAt) {
              isExpired = true;
            } else {
              const last = new Date(metric.lastUpdateAt);
              if (isNaN(last.getTime())) {
                isExpired = true;
              } else {
                const expiryTime = last.getTime() + (metric.refreshInterval || 5) * 60000;
                isExpired = Date.now() >= expiryTime;
              }
            }

            return {
              ...metric,
              lastUpdate: 'Sincronizando...',
              status: isExpired ? 'ok' as const : metric.status,
              value: isExpired ? 0 : metric.value
            };
          }
          return metric;
        });
        return { ...section, metrics: updatedMetrics };
      });
    });

    setEventLog(prev => ([{ 
      id: Math.random().toString(36).substr(2, 9), 
      message: "Sincronização em segundo plano iniciada.", 
      time: new Date().toLocaleTimeString('pt-BR'), 
      type: 'info' as const 
    }, ...prev] as any).slice(0, 50));

    // Fire all SQL queries concurrently (Promise.all-equivalent backend execution but progressive client rendering!)
    metricsToUpdate.forEach(({ metric }, index) => {
      postSqlQuery(metric.sqlQuery!, metric.id)
        .then(res => {
          const resolveNow = new Date();
          const resolveNowIso = resolveNow.toISOString();
          const resolveNowString = resolveNow.toLocaleString('pt-BR');
          const rowCount = Array.isArray(res) ? res.length : 0;
          const isNowError = rowCount > 0;
          const wasError = metric.status === 'error';

          const updatedHistory = [rowCount, ...(metric.history || [])].slice(0, 10);

          // Render card outcome immediately as it completes
          setData(current => {
            return current.map(section => {
              const mIdx = section.metrics.findIndex(m => m.id === metric.id);
              if (mIdx !== -1) {
                const currentMetric = section.metrics[mIdx];
                const updatedMetrics = [...section.metrics];
                updatedMetrics[mIdx] = {
                  ...currentMetric,
                  value: rowCount,
                  details: Array.isArray(res) ? res : [],
                  lastUpdate: resolveNowString,
                  lastUpdateAt: resolveNowIso,
                  status: isNowError ? 'error' : 'ok',
                  history: updatedHistory
                };
                return { ...section, metrics: updatedMetrics };
              }
              return section;
            });
          });

          // Stagger card writes to SharePoint to keep background services fluent and conflict-free!
          setTimeout(() => {
            saveMetricData(metric.id, resolveNowIso, res, updatedHistory)
              .catch(err => console.error(`Error saving metric data for ${metric.id}:`, err));
          }, index * 1000);

          // Specific state alert transition router
          if (isNowError && !wasError) {
            triggerAlarm(`Alerta! Nova divergência detectada na métrica: ${metric.title}`, 'critical');
            try {
              sendDivergenceEmail(metric.title, Array.isArray(res) ? res : []);
            } catch (e1) {
              console.error("Error sending email:", e1);
            }
            try {
              sendTeamsNotification(metric.title, Array.isArray(res) ? res : []);
            } catch (e2) {
              console.error("Error sending Teams notification:", e2);
            }
          } else if (!isNowError && wasError) {
            triggerAlarm(`Excelente! Divergência resolvida na métrica: ${metric.title}`, 'success');
          }
        })
        .catch(err => {
          console.error(`Sincronização do card "${metric.title}" falhou em segundo plano:`, err);
          const resolveNow = new Date();
          setData(current => {
            return current.map(section => {
              const mIdx = section.metrics.findIndex(m => m.id === metric.id);
              if (mIdx !== -1) {
                const currentMetric = section.metrics[mIdx];
                const updatedMetrics = [...section.metrics];
                updatedMetrics[mIdx] = {
                  ...currentMetric,
                  lastUpdate: `Erro de sincronização (${resolveNow.toLocaleTimeString('pt-BR')})`,
                  lastUpdateAt: resolveNow.toISOString()
                };
                return { ...section, metrics: updatedMetrics };
              }
              return section;
            });
          });
        });
    });
  };

  const fetchAllData = () => fetchAllDataInternal(data);

  const refreshSingleMetric = async (metric: Metric) => {
    if (!metric.sqlQuery || isRefreshing) return;
    try {
      const result = await postSqlQuery(metric.sqlQuery, metric.id);
      const now = new Date();
      const nowIso = now.toISOString();
      const rowCount = Array.isArray(result) ? result.length : 0;
      let updatedHistory: number[] = [];
      const wasError = metric.status === 'error';
      const isNowError = rowCount > 0;

      setData(current => {
        return current.map(section => {
          const mIdx = section.metrics.findIndex(m => m.id === metric.id);
          if (mIdx !== -1) {
            const currentMetric = section.metrics[mIdx];
            updatedHistory = [rowCount, ...(currentMetric.history || [])].slice(0, 10);
            const updatedMetrics = [...section.metrics];
            updatedMetrics[mIdx] = { 
              ...currentMetric, 
              value: rowCount, 
              details: Array.isArray(result) ? result : [], 
              lastUpdate: now.toLocaleString('pt-BR'), 
              lastUpdateAt: nowIso, 
              status: rowCount > 0 ? 'error' : 'ok', 
              history: updatedHistory 
            };
            return { ...section, metrics: updatedMetrics };
          }
          return section;
        });
      });
      saveMetricData(metric.id, nowIso, result, updatedHistory);
      setEventLog(prev => ([{ id: Math.random().toString(36).substr(2, 9), message: `Card "${metric.title}" atualizado automaticamente.`, time: now.toLocaleTimeString('pt-BR'), type: 'success' as const }, ...prev] as any).slice(0, 50));
      
      // Trigger notifications if needed
      if (isNowError && !wasError) {
        triggerAlarm(`Alerta! Nova divergência detectada na métrica: ${metric.title}`, 'critical');
        try {
          sendDivergenceEmail(metric.title, Array.isArray(result) ? result : []);
        } catch (e1) {
          console.error("Error triggering email alarm:", e1);
        }

        try {
          sendTeamsNotification(metric.title, Array.isArray(result) ? result : []);
        } catch (e2) {
          console.error("Error triggering Teams alarm:", e2);
        }
      } else if (!isNowError && wasError) {
        triggerAlarm(`Excelente! Divergência resolvida na métrica: ${metric.title}`, 'success');
      }
    } catch (err) { console.error(`Auto-refresh error for metric ${metric.id}:`, err); }
  };

  const totalDivergences = data.reduce((acc, section) => acc + section.metrics.reduce((mAcc, m) => mAcc + Number(m.value || 0), 0), 0);
  const criticalMetrics = data.reduce((acc, section) => acc + section.metrics.filter(m => m.status === 'error').length, 0);
  const allMetrics = data.flatMap(s => s.metrics);
  const totalMetrics = allMetrics.length;
  let totalPoints = 0, totalOkPoints = 0, pointsUnderSLA = 0;
  allMetrics.forEach(m => { if (m.history && m.history.length > 0) { totalPoints += m.history.length; totalOkPoints += m.history.filter(v => v === 0).length; pointsUnderSLA += m.history.filter(v => v <= 10).length; } });
  const accuracyToday = totalPoints > 0 ? (totalOkPoints / totalPoints) * 100 : (totalMetrics > 0 ? (allMetrics.filter(m => m.value === 0).length / totalMetrics) * 100 : 0);
  const slaScore = totalPoints > 0 ? (pointsUnderSLA / totalPoints) * 100 : 100;
  let prevTotalPoints = 0, prevOkPoints = 0;
  allMetrics.forEach(m => { if (m.history && m.history.length > 1) { const h = m.history.slice(1); prevTotalPoints += h.length; prevOkPoints += h.filter(v => v === 0).length; } });
  const currentOkRate = totalMetrics > 0 ? (allMetrics.filter(m => m.value === 0).length / totalMetrics) * 100 : 0;
  const prevOkRate = prevTotalPoints > 0 ? (prevOkPoints / prevTotalPoints) * 100 : currentOkRate;
  const trendValue = currentOkRate - prevOkRate;
  const handleWidthChange = (title: string, width: number) => { setLayoutConfig(prev => prev.map(c => c.title === title ? { ...c, width } : c)); };

  const getFilteredData = () => {
    let result = data;
    
    // 1. Division filter (slug and case-insensitive checking to handle browser search perfectly)
    if (selectedDivision !== 'TODAS') {
      const normalizedSel = normalizeUrlParam(selectedDivision);
      result = result.filter(s => 
        s.title.toUpperCase() === selectedDivision.toUpperCase() || 
        s.title.toLowerCase() === selectedDivision.toLowerCase() ||
        normalizeUrlParam(s.title) === normalizedSel
      );
    }
    
    // 2. Metric status and search box filtering
    return result.map(section => {
      let filteredMetrics = section.metrics;
      
      if (selectedSeverity === 'CRITICO') {
        filteredMetrics = filteredMetrics.filter(m => m.status === 'error');
      } else if (selectedSeverity === 'ESTAVEL') {
        filteredMetrics = filteredMetrics.filter(m => m.status === 'ok');
      }
      
      if (dashboardSearch.trim()) {
        const query = dashboardSearch.toLowerCase();
        filteredMetrics = filteredMetrics.filter(m => 
          m.title.toLowerCase().includes(query) ||
          (m.objective && m.objective.toLowerCase().includes(query))
        );
      }
      
      return {
        ...section,
        metrics: filteredMetrics
      };
    }).filter(section => section.metrics.length > 0); // Hide empty sections matching the criteria
  };

  const getPackedSections = () => {
    const filteredData = getFilteredData();
    if (!filteredData || filteredData.length === 0) return [];
    
    // Map all sections in filteredData to their config if it exists, or a default config
    const sectionsWithConfig = filteredData.map(section => {
      // Try to find config by title (we could use ID in the future if we ensured all sections had one)
      const config = layoutConfig.find(c => c.title === section.title) || { title: section.title, width: 33.33 };
      return { config, section };
    });

    const result: { config: { title: string, width: number }, section: Section }[] = [];
    const remaining = [...sectionsWithConfig];
    
    while (remaining.length > 0) {
      let currentRowWidth = 0;
      const rowIndices: number[] = [];
      
      for (let i = 0; i < remaining.length; i++) { 
        if (currentRowWidth + remaining[i].config.width <= 100.1) { 
          currentRowWidth += remaining[i].config.width; 
          rowIndices.push(i); 
        } 
      }
      
      if (rowIndices.length === 0 && remaining.length > 0) {
        // Current section is too wide for any row, force it as 100% or just add it
        const first = remaining.shift()!;
        result.push(first);
        continue;
      }

      // Map indices to their respective items in the correct order first
      const rowItems = rowIndices.map(idx => remaining[idx]);
      result.push(...rowItems);
      
      // Then remove them in descending index order to avoid index disruption
      rowIndices.sort((a, b) => b - a).forEach(idx => { 
        remaining.splice(idx, 1); 
      });
    }
    return result;
  };

  return (
    <main id="main-dashboard" className={`min-h-screen transition-colors duration-700 p-4 md:p-8 space-y-10 pb-32 ${isWarRoom ? 'bg-[#050510] text-white' : 'bg-slate-50 text-slate-900'}`}>
      {preferences.showHeader && (
        <header id="dashboard-header-print-ignore" className={`flex flex-col sm:flex-row justify-between items-center mb-6 px-6 py-6 rounded-2xl transition-all duration-700 gap-6 mx-auto w-full ${isWarRoom ? 'bg-[#0f1125] border border-indigo-900/40 shadow-[0_0_40px_rgba(0,0,0,0.5)] max-w-[1700px]' : 'bg-white shadow-sm border border-slate-200 max-w-[1400px]'}`}>
          <div className="flex items-center gap-4">
            <div className={`p-3 rounded-2xl transition-colors duration-500 relative ${isWarRoom ? 'bg-indigo-950/30' : 'bg-slate-50'}`}><Activity className={`w-8 h-8 transition-colors duration-500 ${isWarRoom ? 'text-brand-red animate-pulse' : 'text-slate-900'}`} />{isWarRoom && <motion.div animate={{ scale: [1, 1.5, 1], opacity: [0.5, 0, 0.5] }} transition={{ repeat: Infinity, duration: 2 }} className="absolute inset-0 bg-brand-red/20 rounded-2xl" />}</div>
            <div><h1 className={`text-2xl font-black italic tracking-tighter uppercase leading-tight transition-colors duration-500 ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>Monitor <span className="text-brand-red font-black">Operacional</span></h1><div className="flex items-center gap-2 mt-1"><div className={`w-2 h-2 rounded-full animate-pulse ${isWarRoom ? 'bg-red-500 shadow-[0_0_8px_rgba(239,68,68,0.8)]' : 'bg-emerald-500'}`} /><p className={`text-[10px] font-black uppercase tracking-[0.2em] ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Controle de divergências operacionais em tempo real</p></div></div>
          </div>
        </header>
      )}

      {/* KPI Cards Block */}
      {(preferences.showKpiDivergences || preferences.showKpiCritical || preferences.showKpiAccuracy || preferences.showKpiSla) && (
        <div id="dashboard-kpi-block-print-ignore" className={`grid grid-cols-2 md:grid-cols-${
          [preferences.showKpiDivergences, preferences.showKpiCritical, preferences.showKpiAccuracy, preferences.showKpiSla].filter(Boolean).length
        } gap-4 mx-auto w-full transition-all duration-700 ${isWarRoom ? 'max-w-[1700px]' : 'max-w-[1400px]'}`}>
          
          {preferences.showKpiDivergences && (
            <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-100 shadow-sm'}`}>
              <div className="flex justify-between items-start">
                <div>
                  <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Total Divergências</p>
                  <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>{totalDivergences}</h4>
                </div>
                <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-indigo-500/10 text-indigo-400' : 'bg-slate-50 text-slate-400'}`}>
                  <Activity className="w-5 h-5" />
                </div>
              </div>
              <div className={`mt-4 h-1.5 w-full rounded-full overflow-hidden ${isWarRoom ? 'bg-slate-800' : 'bg-slate-100'}`}>
                <motion.div initial={{ width: 0 }} animate={{ width: `${Math.min(100, (totalDivergences / 50) * 100)}%` }} className="h-full bg-brand-red" />
              </div>
            </div>
          )}

          {preferences.showKpiCritical && (
            <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-100 shadow-sm'}`}>
              <div className="flex justify-between items-start">
                <div>
                  <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Métricas Críticas</p>
                  <h4 className="text-3xl font-black italic tracking-tighter text-brand-red">{criticalMetrics}</h4>
                </div>
                <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-red-500/10 text-red-400' : 'bg-red-50 text-red-400'}`}>
                  <XCircle className="w-5 h-5" />
                </div>
              </div>
              <p className="mt-4 text-[10px] font-bold text-slate-500 italic">Requer atenção imediata</p>
            </div>
          )}

          {preferences.showKpiAccuracy && (
            <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${isWarRoom ? 'bg-[#0f1125] border-indigo-900/30' : 'bg-white border-slate-100 shadow-sm'}`}>
              <div className="flex justify-between items-start">
                <div>
                  <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Acertos Hoje</p>
                  <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-emerald-500' : 'text-emerald-600'}`}>{accuracyToday.toFixed(1)}%</h4>
                </div>
                <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-emerald-500/10 text-emerald-400' : 'bg-emerald-50 text-emerald-600'}`}>
                  <CheckCircle2 className="w-5 h-5" />
                </div>
              </div>
              <div className="mt-4 flex items-center gap-1">
                {trendValue >= 0 ? <TrendingUp className="w-3 h-3 text-emerald-500" /> : <TrendingDown className="w-3 h-3 text-red-500" />}
                <span className={`text-[10px] font-black ${trendValue >= 0 ? 'text-emerald-500' : 'text-red-500'}`}>
                  {trendValue >= 0 ? '+' : ''}{trendValue.toFixed(1)}% vs anterior
                </span>
              </div>
            </div>
          )}

          {preferences.showKpiSla && (
            <div className={`p-5 rounded-2xl border transition-all duration-500 overflow-hidden relative group ${
              isWarRoom ? 'bg-brand-red border-red-800 shadow-[0_0_30px_rgba(204,0,0,0.2)]' : 'bg-brand-yellow border-brand-yellow shadow-sm'
            }`}>
              <div className="flex justify-between items-start">
                <div>
                  <p className={`text-[10px] font-black uppercase tracking-widest mb-1 ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>SLA Operacional</p>
                  <h4 className={`text-3xl font-black italic tracking-tighter ${isWarRoom ? 'text-white' : 'text-slate-900'}`}>{slaScore.toFixed(1)}%</h4>
                </div>
                <div className={`p-2 rounded-lg ${isWarRoom ? 'bg-white/10 text-white' : 'bg-white/20 text-slate-900'}`}>
                  <RefreshCcw className="w-5 h-5" />
                </div>
              </div>
              <p className={`mt-4 text-[10px] font-black uppercase ${isWarRoom ? 'text-red-200' : 'text-slate-700'}`}>
                Status: {slaScore > 95 ? 'Excelente' : slaScore > 80 ? 'Bom' : 'Crítico'}
              </p>
            </div>
          )}
        </div>
      )}

      {/* Filtros Operacionais */}
      <div className={`mx-auto w-full p-4 rounded-2xl border transition-all duration-700 ${
        isWarRoom 
          ? 'bg-[#0f1125] border-indigo-900/40 shadow-[0_0_30px_rgba(0,0,0,0.55)] max-w-[1700px] text-white' 
          : 'bg-white border-slate-200 shadow-sm max-w-[1400px] text-slate-900'
      }`}>
        <div className="flex flex-col md:flex-row gap-4 items-stretch md:items-center justify-between">
          
          {/* Busca por Texto */}
          <div className="relative flex-grow max-w-md">
            <span className="absolute inset-y-0 left-0 flex items-center pl-3.5 pointer-events-none text-slate-400">
              <Search className="w-4 h-4" />
            </span>
            <input
              type="text"
              placeholder="Buscar por métrica ou descrição..."
              value={dashboardSearch}
              onChange={(e) => setDashboardSearch(e.target.value)}
              className={`w-full pl-10 pr-10 py-2.5 rounded-xl text-xs outline-none transition-all font-medium border ${
                isWarRoom 
                  ? 'bg-slate-950/40 border-slate-800 placeholder-slate-500 text-slate-100 focus:border-indigo-500 focus:bg-slate-950' 
                  : 'bg-slate-50 border-slate-200 placeholder-slate-400 text-slate-700 focus:border-indigo-600 focus:bg-white focus:ring-2 focus:ring-indigo-100'
              }`}
            />
            {dashboardSearch && (
              <button 
                onClick={() => setDashboardSearch('')} 
                className="absolute inset-y-0 right-0 pr-3 flex items-center text-slate-400 hover:text-slate-600 cursor-pointer"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {/* Opções de Filtro */}
          <div className="flex flex-wrap items-center gap-3.5">
            
            {/* Seletor de Divisão */}
            <div className="flex items-center gap-2">
              <span className={`text-[9px] font-black uppercase tracking-widest ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Divisão:</span>
              <select
                value={selectedDivision}
                onChange={(e) => setSelectedDivision(e.target.value)}
                className={`text-xs font-bold py-2 px-3 rounded-xl border outline-none tracking-tight cursor-pointer transition-all ${
                  isWarRoom
                    ? 'bg-[#181a3a] border-slate-800 text-slate-200 focus:border-indigo-500'
                    : 'bg-slate-50 border-slate-200 text-slate-700 focus:border-indigo-600 focus:bg-white'
                }`}
              >
                <option value="TODAS">TODAS AS DIVISÕES</option>
                {data.map(div => (
                  <option key={div.id || div.title} value={div.title}>{div.title.toUpperCase()}</option>
                ))}
              </select>
            </div>

            {/* Seletor de Severidade / Status */}
            <div className="flex items-center gap-2">
              <span className={`text-[9px] font-black uppercase tracking-widest ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>Status:</span>
              <div className={`flex p-0.5 rounded-lg border ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-100 border-slate-200/50'}`}>
                {(['TODOS', 'CRITICO', 'ESTAVEL'] as const).map((status) => {
                  const label = status === 'TODOS' ? 'Todos' : status === 'CRITICO' ? 'Críticos' : 'Estáveis';
                  const active = selectedSeverity === status;
                  return (
                    <button
                      key={status}
                      onClick={() => setSelectedSeverity(status)}
                      className={`px-3 py-1.5 rounded-md text-[9px] font-black uppercase tracking-wider transition-all cursor-pointer ${
                        active
                          ? isWarRoom 
                            ? 'bg-indigo-600 text-white shadow-sm' 
                            : 'bg-slate-900 text-white shadow-sm'
                          : `text-slate-400 hover:text-slate-200 ${isWarRoom ? 'hover:bg-slate-900/40' : 'hover:bg-slate-200/50 hover:text-slate-700'}`
                      }`}
                    >
                      {label}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Limpar Filtros */}
            {(selectedDivision !== 'TODAS' || selectedSeverity !== 'TODOS' || dashboardSearch) && (
              <button
                onClick={() => {
                  setSelectedDivision('TODAS');
                  setSelectedSeverity('TODOS');
                  setDashboardSearch('');
                }}
                className={`px-3 py-2 rounded-xl text-[9px] font-black uppercase tracking-widest transition-all cursor-pointer flex items-center gap-1.5 ${
                  isWarRoom
                    ? 'bg-red-500/10 text-red-400 hover:bg-red-500/20 border border-red-500/20'
                    : 'bg-rose-50 text-rose-600 hover:bg-rose-100 border border-rose-100'
                }`}
              >
                <Filter className="w-3.5 h-3.5" /> Limpar
              </button>
            )}

          </div>

        </div>
      </div>

      <div className={`flex flex-wrap gap-x-8 gap-y-12 px-2 transition-all duration-700 mx-auto justify-center ${isWarRoom ? 'max-w-[1700px]' : 'max-w-[1400px]'}`}>
        <AnimatePresence mode="popLayout">
          {getPackedSections().map(({ section, config }) => (
            <motion.div layout initial={{ opacity: 0, scale: 0.9 }} animate={{ opacity: 1, scale: 1 }} exit={{ opacity: 0, scale: 0.9 }} key={section.title} className="transition-all duration-500" style={{ width: `calc(${config.width}% - 32px)`, minWidth: '320px', flexGrow: 0, flexShrink: 0 }}>
              <SectionContainer section={section} onCardClick={setSelectedMetric} onCardRefresh={refreshSingleMetric} isWarRoom={isWarRoom} />
            </motion.div>
          ))}
        </AnimatePresence>
      </div>

      <AnimatePresence>{selectedMetric && <DivergenceModal metric={selectedMetric} onClose={() => setSelectedMetric(null)} onRefresh={refreshSingleMetric} enableAI={preferences.enableAI && globalAiEnabled} />}</AnimatePresence>
      <AnimatePresence>{isHelpOpen && <SupportModal onClose={() => setIsHelpOpen(false)} isWarRoom={isWarRoom} />}</AnimatePresence>
      {typeof document !== 'undefined' && createPortal(
        <AnimatePresence>
          {isSettingsOpen && (
            <>
              <motion.div 
                initial={{ opacity: 0 }} 
                animate={{ opacity: 1 }} 
                exit={{ opacity: 0 }} 
                onClick={() => setIsSettingsOpen(false)} 
                className="fixed inset-0 h-screen w-screen bg-slate-950/25 backdrop-blur-[2px] z-[999]" 
              />
              <motion.div 
                initial={{ opacity: 0, x: 300 }} 
                animate={{ opacity: 1, x: 0 }} 
                exit={{ opacity: 0, x: 300 }} 
                className={`fixed right-0 top-0 bottom-0 h-screen w-80 z-[1000] shadow-2xl p-6 border-l transition-colors duration-500 flex flex-col ${isWarRoom ? 'bg-slate-900 border-slate-700 text-white' : 'bg-white border-slate-200 text-slate-900'}`}
              >
                <div className="flex justify-between items-center mb-6"><div className="flex items-center gap-3"><Activity className="w-6 h-6 text-brand-red" /><h2 className="text-xl font-black italic uppercase tracking-tighter">Centro de Controle</h2></div><button onClick={() => setIsSettingsOpen(false)} className={`p-2 rounded-full transition-colors ${isWarRoom ? 'hover:bg-slate-800 text-slate-400' : 'hover:bg-slate-100 text-slate-500'}`}><X className="w-6 h-6" /></button></div>
                <div className={`flex p-1 rounded-xl mb-8 transition-colors ${isWarRoom ? 'bg-slate-950' : 'bg-slate-100'}`}>
                  <button onClick={() => setSettingsTab('layout')} className={`flex-1 py-2 text-[10px] font-black uppercase tracking-widest rounded-lg transition-all ${settingsTab === 'layout' ? 'bg-brand-red text-white shadow-md' : 'text-slate-500 hover:bg-white/10'}`}>Layout</button>
                  <button onClick={() => setSettingsTab('log')} className={`flex-1 py-2 text-[10px] font-black uppercase tracking-widest rounded-lg transition-all ${settingsTab === 'log' ? 'bg-brand-red text-white shadow-md' : 'text-slate-500 hover:bg-white/10'}`}>Atividade</button>
                  <button onClick={() => setSettingsTab('prefs')} className={`flex-1 py-2 text-[10px] font-black uppercase tracking-widest rounded-lg transition-all ${settingsTab === 'prefs' ? 'bg-brand-red text-white shadow-md' : 'text-slate-500 hover:bg-white/10'}`}>Definições</button>
                </div>
                <div className="flex-grow overflow-y-auto pr-2 scrollbar-hide">
                  {settingsTab === 'layout' ? (
                  <div className="space-y-8 animate-in fade-in slide-in-from-right-4 duration-300">
                    {layoutConfig.map((config) => (
                      <div key={config.title} className="space-y-4">
                        <div className="flex justify-between items-end">
                          <p className={`text-[10px] font-black uppercase tracking-widest leading-relaxed max-w-[70%] ${isWarRoom ? 'text-indigo-400' : 'text-slate-400'}`}>
                            {config.title}
                          </p>
                          <span className="text-lg font-black tabular-nums">{config.width}%</span>
                        </div>
                        <input 
                          type="range" 
                          min="30" 
                          max="100" 
                          step="5" 
                          value={config.width} 
                          onChange={(e) => handleWidthChange(config.title, parseInt(e.target.value))} 
                          className={`w-full h-1.5 rounded-lg appearance-none cursor-pointer accent-brand-red ${isWarRoom ? 'bg-slate-800' : 'bg-slate-100'}`} 
                        />
                      </div>
                    ))}
                  </div>
                ) : settingsTab === 'log' ? (
                  <div className="space-y-3 animate-in fade-in slide-in-from-left-4 duration-300">
                    {eventLog.length > 0 ? (
                      eventLog.map((event) => (
                        <div key={event.id} className={`p-3 rounded-xl border-l-4 transition-all ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-200'} ${event.type === 'critical' ? 'border-l-red-500' : event.type === 'success' ? 'border-l-emerald-500' : 'border-l-indigo-500'}`}>
                          <div className="flex justify-between items-center mb-1">
                            <span className={`text-[8px] font-black uppercase tracking-widest ${isWarRoom ? 'text-slate-500' : 'text-slate-400'}`}>{event.time}</span>
                            {event.type === 'critical' && <Triangle className="w-2 h-2 text-red-500 fill-red-500 animate-pulse" />}
                          </div>
                          <p className={`text-[10px] font-bold leading-tight ${isWarRoom ? 'text-slate-300' : 'text-slate-700'}`}>{event.message}</p>
                        </div>
                      ))
                    ) : (
                      <div className="text-center py-20">
                        <p className="text-slate-500 font-bold italic text-xs">Nenhuma atividade recente detectada.</p>
                      </div>
                    )}
                  </div>
                ) : (
                  <div className="space-y-6 animate-in fade-in slide-in-from-right-4 duration-300">
                    {/* AI Config */}
                    <div className={`p-4 rounded-xl border ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-100'} flex flex-col gap-3`}>
                      <div className="flex justify-between items-center">
                        <div>
                          <div className="flex items-center gap-2">
                            <p className="text-xs font-black uppercase tracking-tight">IA Analítica Gemini</p>
                            {!globalAiEnabled && (
                              <span className="text-[8px] font-black uppercase bg-red-100 text-red-600 px-1.5 py-0.5 rounded-full leading-none">
                                Inativo
                              </span>
                            )}
                          </div>
                          <p className="text-[9px] text-slate-500">
                            {globalAiEnabled ? 'Geração de insights e sugestões' : 'Recurso desativado globalmente pelo administrador.'}
                          </p>
                        </div>
                        <button
                          disabled={!globalAiEnabled}
                          onClick={() => savePreferences({ ...preferences, enableAI: !preferences.enableAI })}
                          className={`relative inline-flex h-5 w-10 items-center rounded-full transition-colors duration-300 ${
                            !globalAiEnabled 
                              ? 'bg-slate-200/50 cursor-not-allowed opacity-50' 
                              : preferences.enableAI ? 'bg-brand-red' : isWarRoom ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        >
                          <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform duration-300 ${
                            !globalAiEnabled
                              ? 'translate-x-1'
                              : preferences.enableAI ? 'translate-x-[22px]' : 'translate-x-1'
                          }`} />
                        </button>
                      </div>

                      <div className="border-t border-slate-200/20 my-1" />

                      {/* Header Config */}
                      <div className="flex justify-between items-center">
                        <div>
                          <p className="text-xs font-black uppercase tracking-tight">Exibir Cabeçalho</p>
                          <p className="text-[9px] text-slate-500">Mostrar cabeçalho principal</p>
                        </div>
                        <button
                          onClick={() => savePreferences({ ...preferences, showHeader: !preferences.showHeader })}
                          className={`relative inline-flex h-5 w-10 items-center rounded-full transition-colors duration-300 ${
                            preferences.showHeader ? 'bg-brand-red' : isWarRoom ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        >
                          <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform duration-300 ${
                            preferences.showHeader ? 'translate-x-[22px]' : 'translate-x-1'
                          }`} />
                        </button>
                      </div>
                    </div>

                    {/* KPI Cards Config */}
                    <div className={`p-4 rounded-xl border ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-100'} space-y-4`}>
                      <h3 className="text-[9px] font-black uppercase tracking-wider text-slate-400">Visibilidade dos Cards KPI</h3>
                      
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold">Total Divergências</span>
                        <button
                          onClick={() => savePreferences({ ...preferences, showKpiDivergences: !preferences.showKpiDivergences })}
                          className={`relative inline-flex h-5 w-10 items-center rounded-full transition-colors duration-300 ${
                            preferences.showKpiDivergences ? 'bg-brand-red' : isWarRoom ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        >
                          <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform duration-300 ${
                            preferences.showKpiDivergences ? 'translate-x-[22px]' : 'translate-x-1'
                          }`} />
                        </button>
                      </div>

                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold">Métricas Críticas</span>
                        <button
                          onClick={() => savePreferences({ ...preferences, showKpiCritical: !preferences.showKpiCritical })}
                          className={`relative inline-flex h-5 w-10 items-center rounded-full transition-colors duration-300 ${
                            preferences.showKpiCritical ? 'bg-brand-red' : isWarRoom ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        >
                          <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform duration-300 ${
                            preferences.showKpiCritical ? 'translate-x-[22px]' : 'translate-x-1'
                          }`} />
                        </button>
                      </div>

                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold">Acertos Hoje</span>
                        <button
                          onClick={() => savePreferences({ ...preferences, showKpiAccuracy: !preferences.showKpiAccuracy })}
                          className={`relative inline-flex h-5 w-10 items-center rounded-full transition-colors duration-300 ${
                            preferences.showKpiAccuracy ? 'bg-brand-red' : isWarRoom ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        >
                          <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform duration-300 ${
                            preferences.showKpiAccuracy ? 'translate-x-[22px]' : 'translate-x-1'
                          }`} />
                        </button>
                      </div>

                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold">SLA Operacional</span>
                        <button
                          onClick={() => savePreferences({ ...preferences, showKpiSla: !preferences.showKpiSla })}
                          className={`relative inline-flex h-5 w-10 items-center rounded-full transition-colors duration-300 ${
                            preferences.showKpiSla ? 'bg-brand-red' : isWarRoom ? 'bg-slate-800' : 'bg-slate-200'
                          }`}
                        >
                          <span className={`inline-block h-3.5 w-3.5 transform rounded-full bg-white transition-transform duration-300 ${
                            preferences.showKpiSla ? 'translate-x-[22px]' : 'translate-x-1'
                          }`} />
                        </button>
                      </div>
                    </div>

                    {/* Audio Config */}
                    <div className={`p-4 rounded-xl border ${isWarRoom ? 'bg-slate-950/40 border-slate-800' : 'bg-slate-50 border-slate-100'} space-y-4`}>
                      <div>
                        <h3 className="text-[9px] font-black uppercase tracking-wider text-slate-400">Bips e Alertas de Voz</h3>
                        <p className="text-[8px] text-slate-500">Notificação instantânea para novos desvios</p>
                      </div>
                      
                      <div className="grid grid-cols-2 gap-1.5">
                        {[
                          { key: 'none', label: 'Desativado' },
                          { key: 'beep', label: 'Bipes' },
                          { key: 'ts', label: 'Voz (Autofalante)' },
                          { key: 'both', label: 'Bipe + Voz' },
                        ].map(opt => (
                          <button
                            key={opt.key}
                            onClick={() => {
                              savePreferences({ ...preferences, audioAlertMode: opt.key as any });
                            }}
                            className={`py-1.5 px-2 rounded-lg text-[9px] font-black uppercase tracking-tight transition-all border text-center cursor-pointer ${
                              preferences.audioAlertMode === opt.key
                                ? 'bg-brand-red border-brand-red text-white shadow-sm'
                                : isWarRoom
                                ? 'bg-slate-950 border-slate-800 text-slate-400 hover:text-white'
                                : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-100'
                            }`}
                          >
                            {opt.label}
                          </button>
                        ))}
                      </div>

                      {/* Test sound board */}
                      <div className="pt-3 border-t border-slate-200/10 space-y-2">
                        <span className="text-[8px] font-black uppercase tracking-widest text-slate-400 flex items-center gap-1.5">
                          <Volume2 className="w-3.5 h-3.5 text-brand-red" /> Testar Sons do Sistema
                        </span>
                        <div className="grid grid-cols-1 gap-1.5">
                          <button
                            onClick={() => playAlertBeep('critical')}
                            className={`w-full text-[9px] font-black uppercase tracking-tight py-1.5 rounded border border-transparent transition-all text-left px-3 flex justify-between items-center cursor-pointer ${
                              isWarRoom ? 'bg-indigo-950/30 hover:bg-indigo-950/60 text-slate-300 hover:text-white' : 'bg-slate-100/80 hover:bg-slate-200 text-slate-700'
                            }`}
                          >
                            <span className="flex items-center gap-1">🔊 Alerta de Divergência</span>
                            <span className="text-brand-red font-black text-[8px]">Crítico</span>
                          </button>
                          <button
                            onClick={() => playAlertBeep('success')}
                            className={`w-full text-[9px] font-black uppercase tracking-tight py-1.5 rounded border border-transparent transition-all text-left px-3 flex justify-between items-center cursor-pointer ${
                              isWarRoom ? 'bg-indigo-950/30 hover:bg-indigo-950/60 text-slate-300 hover:text-white' : 'bg-slate-100/80 hover:bg-slate-200 text-slate-700'
                            }`}
                          >
                            <span className="flex items-center gap-1">🔊 Sucesso / Resolução</span>
                            <span className="text-emerald-500 font-black text-[8px]">Sucesso</span>
                          </button>
                          <button
                            onClick={() => speakAlertText("Mensagem de teste do auto-falante: Operação de monitor normalizada.")}
                            className={`w-full text-[9px] font-black uppercase tracking-tight py-1.5 rounded border border-transparent transition-all text-left px-3 flex justify-between items-center cursor-pointer ${
                              isWarRoom ? 'bg-indigo-950/30 hover:bg-indigo-950/60 text-slate-300 hover:text-white' : 'bg-slate-100/80 hover:bg-slate-200 text-slate-700'
                            }`}
                          >
                            <span className="flex items-center gap-1.5">🗣️ Sintetizador de Voz</span>
                            <span className="text-indigo-400 font-black text-[8px]">pt-BR</span>
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
              <div className={`mt-auto pt-6 p-4 rounded-xl border italic text-[10px] font-medium leading-relaxed ${isWarRoom ? 'bg-slate-950/50 border-slate-800 text-slate-500' : 'bg-slate-50 border-slate-100 text-slate-400'}`}>
                {settingsTab === 'layout' 
                  ? "Sincronização inteligente de grid baseada em larguras customizadas." 
                  : settingsTab === 'log' 
                  ? "Log operacional de mutações detectadas em tempo real via stream."
                  : "Preferências de exibição, recursos de IA analítica e notificações sonoras."}
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>,
      document.body
    )}

      <button 
        onClick={() => setIsHelpOpen(true)} 
        className={`fixed bottom-[168px] right-8 z-50 h-14 w-14 hover:w-[220px] rounded-2xl shadow-2xl flex items-center p-4 transition-all duration-300 ease-in-out active:scale-95 group overflow-hidden ${isWarRoom ? 'bg-brand-red text-white hover:bg-red-600 shadow-[0_10px_30px_rgba(204,0,0,0.5)]' : 'bg-[#0005a3] text-white hover:bg-blue-900 shadow-[0_10px_30px_rgba(0,5,163,0.3)]'}`}
      >
        <HelpCircle className="w-6 h-6 flex-shrink-0" />
        <span className="font-black text-xs uppercase tracking-widest border-l border-white/20 pl-3 ml-0 group-hover:ml-3 whitespace-nowrap opacity-0 group-hover:opacity-100 transition-all duration-300 pointer-events-none max-w-0 group-hover:max-w-xs overflow-hidden">
          Suporte e Ajuda
        </span>
      </button>

      <button 
        onClick={() => setIsWarRoom(!isWarRoom)} 
        className={`fixed bottom-[100px] right-8 z-50 h-14 w-14 hover:w-[250px] rounded-2xl shadow-2xl flex items-center p-4 transition-all duration-300 ease-in-out active:scale-95 group overflow-hidden ${isWarRoom ? 'bg-brand-red text-white hover:bg-red-600 shadow-[0_10px_30px_rgba(204,0,0,0.5)]' : 'bg-[#0f1125] text-white hover:bg-slate-900 shadow-[0_10px_30px_rgba(0,0,0,0.3)]'}`}
      >
        <Activity className={`w-6 h-6 flex-shrink-0 ${isWarRoom ? 'animate-pulse text-white' : 'text-brand-red'}`} />
        <span className="font-black text-xs uppercase tracking-widest border-l border-white/20 pl-3 ml-0 group-hover:ml-3 whitespace-nowrap opacity-0 group-hover:opacity-100 transition-all duration-300 pointer-events-none max-w-0 group-hover:max-w-xs overflow-hidden">
          {isWarRoom ? 'Sair do Modo War Room' : 'Ativar Modo War Room'}
        </span>
      </button>

      <button 
        onClick={() => setIsSettingsOpen(true)} 
        className={`fixed bottom-8 right-8 z-50 h-14 w-14 hover:w-[220px] rounded-2xl shadow-2xl flex items-center p-4 transition-all duration-300 ease-in-out active:scale-95 group overflow-hidden ${isWarRoom ? 'bg-brand-red text-white hover:bg-red-600 shadow-[0_10px_30px_rgba(204,0,0,0.5)]' : 'bg-slate-900 text-white hover:bg-slate-800 shadow-[0_10px_30px_rgba(0,0,0,0.3)]'}`}
      >
        <Settings className="w-6 h-6 flex-shrink-0 group-hover:rotate-180 transition-transform duration-700" />
        <span className="font-black text-xs uppercase tracking-widest border-l border-white/20 pl-3 ml-0 group-hover:ml-3 whitespace-nowrap opacity-0 group-hover:opacity-100 transition-all duration-300 pointer-events-none max-w-0 group-hover:max-w-xs overflow-hidden">
          Ajustar Painel
        </span>
      </button>
      <footer className={`text-center text-[10px] py-10 border-t mt-12 transition-colors duration-500 mx-auto w-full ${isWarRoom ? 'border-indigo-950 text-indigo-900 max-w-[1700px]' : 'border-slate-200 text-slate-400 max-w-[1400px]'}`}>&copy; {new Date().getFullYear()} Monitoring System - Todos os direitos reservados</footer>
    </main>
  );
}
