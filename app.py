import React, { useState, useRef, useEffect } from 'react';
import { 
  UploadCloud, 
  FileText, 
  Settings, 
  CheckCircle, 
  Download, 
  RefreshCw, 
  FileCheck,
  AlignLeft,
  Layout,
  ChevronRight,
  ShieldCheck,
  AlertCircle
} from 'lucide-react';

export default function TranslationPlatform() {
  const [step, setStep] = useState(1); // 1: Upload, 2: Config, 3: Processing, 4: Done
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  
  // Configuration State
  const [formatType, setFormatType] = useState('simples'); // 'simples' or 'juramentada'
  const [translatorName, setTranslatorName] = useState('');
  const [registrationNumber, setRegistrationNumber] = useState('');
  
  // Processing State
  const [progress, setProgress] = useState(0);
  const [processingStatus, setProcessingStatus] = useState('');

  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  };

  const handleFile = (uploadedFile) => {
    if (uploadedFile.type !== 'application/pdf') {
      // In a real app, use a toast. Here we just gently alert via UI.
      alert('Por favor, envie apenas arquivos PDF.');
      return;
    }
    setFile(uploadedFile);
    setStep(2);
  };

  const startProcessing = () => {
    setStep(3);
    setProgress(0);
    
    const steps = [
      { progress: 10, msg: 'Lendo arquivo PDF e extraindo páginas...' },
      { progress: 30, msg: 'Aplicando OCR e reconhecendo estrutura...' },
      { progress: 50, msg: formatType === 'juramentada' ? 'Identificando carimbos e assinaturas...' : 'Mapeando layout, imagens e tabelas...' },
      { progress: 70, msg: formatType === 'juramentada' ? 'Aplicando formatação padrão da junta comercial...' : 'Recriando espelho do documento original...' },
      { progress: 90, msg: 'Revisando formatação final com IA...' },
      { progress: 100, msg: 'Gerando arquivo Word (.docx)...' }
    ];

    let currentStep = 0;

    const interval = setInterval(() => {
      if (currentStep < steps.length) {
        setProgress(steps[currentStep].progress);
        setProcessingStatus(steps[currentStep].msg);
        currentStep++;
      } else {
        clearInterval(interval);
        setTimeout(() => setStep(4), 500);
      }
    }, 1200); // Simulate time for each step
  };

  const handleDownload = () => {
    // We create a mock HTML string that MS Word can interpret as a document
    const content = formatType === 'juramentada' 
      ? `<html><head><meta charset='utf-8'></head><body style="font-family: Courier New, monospace;">
          <h2 style="text-align: center;">TRADUÇÃO PÚBLICA JURAMENTADA</h2>
          <p><strong>Tradutor:</strong> ${translatorName || 'Não informado'}</p>
          <p><strong>Matrícula:</strong> ${registrationNumber || 'Não informado'}</p>
          <hr/>
          <p>[Início da Tradução]</p>
          <p>Este é um documento de exemplo gerado pela plataforma de IA.</p>
          <p>[Carimbo ilegível]</p>
          <p>[Fim da Tradução]</p>
         </body></html>`
      : `<html><head><meta charset='utf-8'></head><body style="font-family: Arial, sans-serif;">
          <div style="border: 1px solid #ccc; padding: 20px;">
            <h1>Documento Formatado (Simples)</h1>
            <p>Este layout tenta imitar o documento original perfeitamente.</p>
            <table border="1" width="100%"><tr><td>Dado 1</td><td>Dado 2</td></tr></table>
          </div>
         </body></html>`;

    const blob = new Blob(['\ufeff', content], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Documento_Formatado_${formatType}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const resetApp = () => {
    setStep(1);
    setFile(null);
    setFormatType('simples');
    setTranslatorName('');
    setRegistrationNumber('');
    setProgress(0);
  };

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800 font-sans p-4 md:p-8 flex flex-col items-center">
      
      {/* Header */}
      <header className="w-full max-w-4xl flex items-center justify-between mb-8 pb-4 border-b border-slate-200">
        <div className="flex items-center gap-2">
          <div className="bg-blue-600 p-2 rounded-lg">
            <RefreshCw className="w-6 h-6 text-white" />
          </div>
          <div>
            <h1 className="text-2xl font-bold text-slate-900 leading-tight">TranslateFlow AI</h1>
            <p className="text-sm text-slate-500">Automação de Formatação Simples e Juramentada</p>
          </div>
        </div>
      </header>

      {/* Main Content Area */}
      <main className="w-full max-w-4xl bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden flex flex-col md:flex-row min-h-[500px]">
        
        {/* Sidebar / Stepper */}
        <div className="bg-slate-100 w-full md:w-64 p-6 border-b md:border-b-0 md:border-r border-slate-200 flex flex-col gap-6">
          <StepIndicator currentStep={step} stepNum={1} icon={<UploadCloud size={20} />} title="Upload do PDF" />
          <StepIndicator currentStep={step} stepNum={2} icon={<Settings size={20} />} title="Configuração" />
          <StepIndicator currentStep={step} stepNum={3} icon={<RefreshCw size={20} />} title="Processamento IA" />
          <StepIndicator currentStep={step} stepNum={4} icon={<Download size={20} />} title="Resultado Word" />
        </div>

        {/* Dynamic Content Area */}
        <div className="flex-1 p-6 md:p-10 flex flex-col justify-center relative">
          
          {/* STEP 1: UPLOAD */}
          {step === 1 && (
            <div className="w-full max-w-md mx-auto animate-fade-in text-center">
              <h2 className="text-2xl font-semibold mb-2">Importar Documento</h2>
              <p className="text-slate-500 mb-8">Envie o documento PDF original que precisa de formatação.</p>
              
              <div 
                className={`border-2 border-dashed rounded-xl p-10 transition-colors cursor-pointer flex flex-col items-center justify-center gap-4
                  ${isDragging ? 'border-blue-500 bg-blue-50' : 'border-slate-300 hover:border-blue-400 hover:bg-slate-50'}`}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current.click()}
              >
                <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center text-blue-600 mb-2">
                  <FileText size={32} />
                </div>
                <div>
                  <p className="font-medium text-slate-700">Clique para buscar ou arraste o arquivo</p>
                  <p className="text-sm text-slate-400 mt-1">Apenas arquivos PDF (Max 20MB)</p>
                </div>
                <input 
                  type="file" 
                  ref={fileInputRef} 
                  onChange={handleFileChange} 
                  accept="application/pdf" 
                  className="hidden" 
                />
              </div>
            </div>
          )}

          {}
          {/* STEP 2: CONFIGURATION */}
          {step === 2 && (
            <div className="w-full max-w-lg mx-auto animate-fade-in">
              <div className="flex items-center gap-3 mb-6 bg-blue-50 p-3 rounded-lg border border-blue-100">
                <FileCheck className="text-blue-600" />
                <span className="font-medium text-blue-900 truncate">{file?.name}</span>
                <span className="text-sm text-blue-500 ml-auto bg-blue-100 px-2 py-1 rounded">{(file?.size / 1024 / 1024).toFixed(2)} MB</span>
              </div>

              <h2 className="text-2xl font-semibold mb-6">Tipo de Formatação</h2>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-8">
                <button
                  onClick={() => setFormatType('simples')}
                  className={`p-4 border-2 rounded-xl text-left transition-all flex flex-col gap-2
                    ${formatType === 'simples' ? 'border-blue-500 bg-blue-50 shadow-sm' : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg w-fit ${formatType === 'simples' ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-600'}`}>
                    <Layout size={20} />
                  </div>
                  <h3 className="font-semibold text-lg mt-1">Formatação Simples</h3>
                  <p className="text-sm text-slate-500 leading-snug">Layout idêntico ao original. Mantém tabelas, imagens e colunas exatas.</p>
                </button>

                <button
                  onClick={() => setFormatType('juramentada')}
                  className={`p-4 border-2 rounded-xl text-left transition-all flex flex-col gap-2
                    ${formatType === 'juramentada' ? 'border-blue-500 bg-blue-50 shadow-sm' : 'border-slate-200 hover:border-slate-300 hover:bg-slate-50'}`}
                >
                  <div className={`p-2 rounded-lg w-fit ${formatType === 'juramentada' ? 'bg-blue-600 text-white' : 'bg-slate-200 text-slate-600'}`}>
                    <ShieldCheck size={20} />
                  </div>
                  <h3 className="font-semibold text-lg mt-1">Juramentada</h3>
                  <p className="text-sm text-slate-500 leading-snug">Padrão oficial. Transcreve imagens como texto, aplica cabeçalho e rodapé.</p>
                </button>
              </div>

              {formatType === 'juramentada' && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-5 mb-8 animate-fade-in space-y-4">
                  <h4 className="font-medium text-slate-700 mb-2 flex items-center gap-2">
                    <AlignLeft size={16} /> Configurações do Tradutor
                  </h4>
                  <div>
                    <label className="block text-sm text-slate-600 mb-1">Nome do Tradutor Público</label>
                    <input 
                      type="text" 
                      value={translatorName}
                      onChange={(e) => setTranslatorName(e.target.value)}
                      placeholder="Ex: João Silva"
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                  <div>
                    <label className="block text-sm text-slate-600 mb-1">Matrícula / Junta Comercial</label>
                    <input 
                      type="text" 
                      value={registrationNumber}
                      onChange={(e) => setRegistrationNumber(e.target.value)}
                      placeholder="Ex: JUCESP Nº 12345"
                      className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                    />
                  </div>
                </div>
              )}

              <div className="flex justify-between items-center mt-auto">
                <button 
                  onClick={() => setStep(1)}
                  className="px-4 py-2 text-slate-600 hover:bg-slate-100 rounded-lg transition-colors font-medium"
                >
                  Voltar
                </button>
                <button 
                  onClick={startProcessing}
                  className="px-6 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors font-medium flex items-center gap-2"
                >
                  Processar Documento <ChevronRight size={18} />
                </button>
              </div>
            </div>
          )}

          {}
          {/* STEP 3: PROCESSING */}
          {step === 3 && (
            <div className="w-full max-w-md mx-auto text-center animate-fade-in py-10">
              <div className="relative w-24 h-24 mx-auto mb-8">
                <svg className="animate-spin w-full h-full text-blue-100" viewBox="0 0 100 100">
                  <circle cx="50" cy="50" r="45" fill="none" strokeWidth="8" stroke="currentColor" />
                </svg>
                <svg className="animate-spin w-full h-full text-blue-600 absolute top-0 left-0" viewBox="0 0 100 100" style={{ animationDirection: 'reverse', animationDuration: '3s' }}>
                  <circle cx="50" cy="50" r="45" fill="none" strokeWidth="8" stroke="currentColor" strokeDasharray="283" strokeDashoffset={283 - (283 * progress) / 100} className="transition-all duration-300" />
                </svg>
                <div className="absolute top-0 left-0 w-full h-full flex items-center justify-center font-bold text-xl text-blue-600">
                  {progress}%
                </div>
              </div>
              
              <h2 className="text-2xl font-semibold mb-2">A IA está trabalhando...</h2>
              <p className="text-slate-500 font-medium h-6">{processingStatus}</p>
              
              <div className="mt-8 bg-slate-50 p-4 rounded-lg border border-slate-200 text-left">
                <p className="text-xs text-slate-500 uppercase tracking-wider font-semibold mb-2">Log do Sistema</p>
                <div className="space-y-2 text-sm text-slate-600 font-mono">
                  <p className={progress >= 10 ? 'opacity-100 text-green-600' : 'opacity-40'}>[✓] Leitura inicial do PDF</p>
                  <p className={progress >= 30 ? 'opacity-100 text-green-600' : 'opacity-40'}>[✓] Processamento de OCR Avançado</p>
                  <p className={progress >= 50 ? 'opacity-100 text-green-600' : 'opacity-40'}>[✓] Mapeamento de Layout / Componentes</p>
                  <p className={progress >= 90 ? 'opacity-100 text-green-600' : 'opacity-40'}>[✓] Geração de Estrutura Word</p>
                </div>
              </div>
            </div>
          )}

          {}
          {/* STEP 4: RESULT */}
          {step === 4 && (
            <div className="w-full max-w-md mx-auto text-center animate-fade-in">
              <div className="w-20 h-20 bg-green-100 text-green-600 rounded-full flex items-center justify-center mx-auto mb-6">
                <CheckCircle size={40} />
              </div>
              
              <h2 className="text-2xl font-semibold mb-2">Documento Pronto!</h2>
              <p className="text-slate-500 mb-8">
                A formatação {formatType === 'juramentada' ? 'juramentada' : 'simples'} foi aplicada com sucesso. O arquivo editável já pode ser baixado.
              </p>

              <div className="bg-blue-50 border border-blue-200 rounded-xl p-6 mb-8">
                <div className="flex items-center gap-4 mb-4">
                  <div className="bg-white p-3 rounded-lg shadow-sm">
                    <FileText className="text-blue-600" size={32} />
                  </div>
                  <div className="text-left flex-1">
                    <h4 className="font-semibold text-slate-800 line-clamp-1">{file?.name.replace('.pdf', '')}_FORMATADO.docx</h4>
                    <p className="text-sm text-slate-500">Pronto para edição e revisão</p>
                  </div>
                </div>
                
                <button 
                  onClick={handleDownload}
                  className="w-full py-3 bg-blue-600 hover:bg-blue-700 text-white rounded-lg transition-colors font-medium flex items-center justify-center gap-2 shadow-sm shadow-blue-200"
                >
                  <Download size={20} /> Baixar Documento Word
                </button>
              </div>

              <div className="flex items-center justify-center gap-2 text-slate-500 text-sm">
                <AlertCircle size={16} />
                <span>Recomendamos sempre uma revisão humana final.</span>
              </div>

              <button 
                onClick={resetApp}
                className="mt-8 text-slate-400 hover:text-slate-600 underline font-medium text-sm transition-colors"
              >
                Formatar outro documento
              </button>
            </div>
          )}

        </div>
      </main>
    </div>
  );
}

function StepIndicator({ currentStep, stepNum, icon, title }) {
  const isCompleted = currentStep > stepNum;
  const isActive = currentStep === stepNum;
  const isPending = currentStep < stepNum;

  return (
    <div className={`flex items-center gap-3 transition-opacity duration-300 ${isPending ? 'opacity-40' : 'opacity-100'}`}>
      <div className={`w-10 h-10 rounded-full flex items-center justify-center shrink-0 border-2 transition-colors
        ${isCompleted ? 'bg-green-500 border-green-500 text-white' : ''}
        ${isActive ? 'border-blue-600 text-blue-600 bg-blue-50' : ''}
        ${isPending ? 'border-slate-300 text-slate-400' : ''}
      `}>
        {isCompleted ? <CheckCircle size={20} /> : icon}
      </div>
      <div className="hidden md:block">
        <p className={`font-semibold text-sm ${isActive ? 'text-blue-600' : 'text-slate-700'}`}>Passo {stepNum}</p>
        <p className="text-xs text-slate-500">{title}</p>
      </div>
    </div>
  );
}
