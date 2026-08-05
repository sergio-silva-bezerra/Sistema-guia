import React, { useState } from 'react';
import { FileText, Download, CheckCircle, ShieldCheck, Briefcase, HardDrive, UserCheck } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { downloadRequirementsDoc, downloadStrategicPlanDoc, downloadFrontendFuncDoc } from '../utils/docGenerator';

interface DocDownloadModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const DocDownloadModal: React.FC<DocDownloadModalProps> = ({ isOpen, onClose }) => {
  const [downloadedType, setDownloadedType] = useState<'req' | 'strat' | 'front' | null>(null);

  const handleDownloadFrontendFunc = () => {
    downloadFrontendFuncDoc();
    setDownloadedType('front');
    setTimeout(() => {
      setDownloadedType(null);
    }, 4000);
  };

  const handleDownloadRequirements = () => {
    downloadRequirementsDoc();
    setDownloadedType('req');
    setTimeout(() => {
      setDownloadedType(null);
    }, 4000);
  };

  const handleDownloadStrategicPlan = () => {
    downloadStrategicPlanDoc();
    setDownloadedType('strat');
    setTimeout(() => {
      setDownloadedType(null);
    }, 4000);
  };

  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm">
        <motion.div
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 10 }}
          className="bg-[var(--bg-secondary)] border border-[var(--border-color)] rounded-xl max-w-lg w-full p-6 shadow-2xl relative overflow-hidden"
        >
          {/* Header */}
          <div className="flex items-start justify-between gap-4 mb-5 pb-4 border-b border-[var(--border-color)]">
            <div className="flex items-center gap-3">
              <div className="p-3 bg-blue-600/10 text-blue-600 dark:text-blue-400 rounded-lg border border-blue-500/20">
                <FileText className="w-6 h-6" />
              </div>
              <div>
                <h3 className="font-bold text-base text-[var(--text-primary)] uppercase tracking-tight">
                  Central de Documentos (.DOC)
                </h3>
                <p className="text-xs text-[var(--text-secondary)] font-medium">
                  Baixe a documentação oficial para Microsoft Word e Google Docs
                </p>
              </div>
            </div>
            <button
              onClick={onClose}
              className="text-[var(--text-secondary)] hover:text-[var(--text-primary)] text-lg font-bold p-1 rounded-md transition-colors"
            >
              ✕
            </button>
          </div>

          {/* Highlights summary */}
          <div className="space-y-3 mb-6 text-xs text-[var(--text-secondary)]">
            <div className="p-3 bg-[var(--bg-tertiary)] rounded-lg border border-[var(--border-color)] flex items-center gap-2.5">
              <ShieldCheck className="w-4 h-4 text-emerald-500 shrink-0" />
              <span><strong>Formatos Compativeis:</strong> Documentos gerados nativamente em .DOC para Word, LibreOffice e Docs.</span>
            </div>

            <div className="p-3 bg-[var(--bg-tertiary)] rounded-lg border border-[var(--border-color)] flex items-center gap-2.5">
              <Briefcase className="w-4 h-4 text-blue-500 shrink-0" />
              <span><strong>Plano Estratégico de Negócio:</strong> Análise B2G/B2Pol, modelo de precificação, Roadmap 2026 e Análise SWOT/LGPD.</span>
            </div>

            <div className="p-3 bg-[var(--bg-tertiary)] rounded-lg border border-[var(--border-color)] flex items-center gap-2.5">
              <HardDrive className="w-4 h-4 text-amber-500 shrink-0" />
              <span><strong>Especificação de Requisitos:</strong> Requisitos Funcionais (RF01-RF12), Não-Funcionais (RNF01-RNF05) e Arquitetura.</span>
            </div>
          </div>

          {/* Action buttons */}
          <div className="flex flex-col gap-2.5">
            <button
              onClick={handleDownloadFrontendFunc}
              className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3.5 px-4 rounded-lg flex items-center justify-center gap-2 transition-all shadow-md active:scale-95 text-xs uppercase tracking-wider"
            >
              {downloadedType === 'front' ? (
                <>
                  <CheckCircle className="w-4 h-4 text-emerald-300" />
                  <span>Manual do Front-End Baixado (.DOC)!</span>
                </>
              ) : (
                <>
                  <Download className="w-4 h-4" />
                  <span>Baixar Manual do Front-End &amp; Perfis (.DOC)</span>
                </>
              )}
            </button>

            <button
              onClick={handleDownloadStrategicPlan}
              className="w-full bg-[var(--bg-tertiary)] hover:bg-[var(--bg-primary)] text-[var(--text-primary)] font-bold py-3 px-4 rounded-lg border border-[var(--border-color)] flex items-center justify-center gap-2 transition-all active:scale-95 text-xs uppercase tracking-wider"
            >
              {downloadedType === 'strat' ? (
                <>
                  <CheckCircle className="w-4 h-4 text-emerald-500" />
                  <span>Plano Estratégico Baixado (.DOC)!</span>
                </>
              ) : (
                <>
                  <Briefcase className="w-4 h-4 text-blue-500" />
                  <span>Baixar Plano Estratégico (.DOC)</span>
                </>
              )}
            </button>

            <button
              onClick={handleDownloadRequirements}
              className="w-full bg-[var(--bg-tertiary)] hover:bg-[var(--bg-primary)] text-[var(--text-primary)] font-bold py-3 px-4 rounded-lg border border-[var(--border-color)] flex items-center justify-center gap-2 transition-all active:scale-95 text-xs uppercase tracking-wider"
            >
              {downloadedType === 'req' ? (
                <>
                  <CheckCircle className="w-4 h-4 text-emerald-500" />
                  <span>Requisitos Baixados (.DOC)!</span>
                </>
              ) : (
                <>
                  <Download className="w-4 h-4 text-blue-500" />
                  <span>Baixar Requisitos &amp; Arquitetura (.DOC)</span>
                </>
              )}
            </button>

            <button
              onClick={onClose}
              className="mt-1 w-full bg-transparent hover:bg-[var(--bg-tertiary)] text-[var(--text-secondary)] font-bold py-2 px-4 rounded-lg transition-colors text-xs uppercase tracking-wider"
            >
              Fechar
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};

