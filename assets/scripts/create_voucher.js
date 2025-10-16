/**
 * VoucherPro - Gerador de Vouchers em Word
 * Script para gerar vouchers de hospedagem em formato .docx
 * usando os dados do formulário HTML
 */

class VoucherGenerator {
    constructor() {
        this.templateUrl = '../assets/templates/voucher_template.docx';
        this.isGenerating = false;
        this.init();
    }

    /**
     * Inicializa o gerador de vouchers
     */
    init() {
        this.bindEvents();
        console.log('VoucherGenerator inicializado');
    }

    /**
     * Vincula eventos aos elementos do formulário
     */
    bindEvents() {
        const confirmButton = document.getElementById('confirmButton');
        if (confirmButton) {
            confirmButton.addEventListener('click', (e) => {
                e.preventDefault();
                this.handleVoucherCreation();
            });
        }
    }

    /**
     * Manipula a criação do voucher
     */
    async handleVoucherCreation() {
        if (this.isGenerating) return;

        const form = document.getElementById('voucherForm');
        if (!form.checkValidity()) {
            form.reportValidity();
            return;
        }

        if (!this.validateDates()) {
            return;
        }

        try {
            this.isGenerating = true;
            this.showLoadingState();
            
            // Coleta dados do formulário
            const voucherData = this.collectFormData();
            
            // Gera o voucher
            await this.generateVoucher(voucherData);
            
            // Mostra modal de sucesso
            this.showSuccessModal();
            
        } catch (error) {
            console.error('Erro ao gerar voucher:', error);
            this.showError('Erro ao gerar voucher. Tente novamente.');
        } finally {
            this.isGenerating = false;
            this.hideLoadingState();
        }
    }

    /**
     * Coleta dados do formulário
     * @returns {Object} Dados do voucher formatados
     */
    collectFormData() {
        const form = document.getElementById('voucherForm');
        const formData = new FormData(form);
        
        // Coleta dados básicos
        const guestName = document.getElementById('guestName').value;
        const accommodationType = document.getElementById('accommodationType').value;
        const observations = document.getElementById('observations').value;
        const checkinDate = document.getElementById('checkinDate').value;
        const checkoutDate = document.getElementById('checkoutDate').value;
        const dailyRate = parseFloat(document.getElementById('dailyRate').value) || 0;
        const advancePayment = parseFloat(document.getElementById('advancePayment').value) || 0;

        // Calcula valores
        const nights = this.calculateNights(checkinDate, checkoutDate);
        const totalValue = nights * dailyRate;
        const remainingValue = Math.max(0, totalValue - advancePayment);

        // Formata datas
        const formattedCheckin = this.formatDate(checkinDate);
        const formattedCheckout = this.formatDate(checkoutDate);
        const currentDate = this.formatDate(new Date().toISOString().split('T')[0]);

        // Formata valores monetários
        const formattedDailyRate = this.formatCurrency(dailyRate);
        const formattedAdvancePayment = this.formatCurrency(advancePayment);
        const formattedTotalValue = this.formatCurrency(totalValue);
        const formattedRemainingValue = this.formatCurrency(remainingValue);

        // Mapeia tipo de acomodação
        const accommodationTypes = {
            'single': 'Quarto Individual',
            'double': 'Quarto Duplo',
            'twin': 'Quarto Twin',
            'suite': 'Suíte',
            'family': 'Quarto Familiar'
        };

        return {
            // Dados básicos
            nome: guestName,
            tipoAcomodacao: accommodationTypes[accommodationType] || accommodationType,
            observacoes: observations || 'Nenhuma observação adicional.',
            
            // Datas
            checkin: formattedCheckin,
            checkout: formattedCheckout,
            dataEmissao: currentDate,
            
            // Valores
            valorDiaria: formattedDailyRate,
            valorAntecipado: formattedAdvancePayment,
            total: formattedTotalValue,
            restante: formattedRemainingValue,
            quantidadeDiarias: nights,
            
            // Dados adicionais
            voucherId: this.generateVoucherId(),
            status: remainingValue > 0 ? 'Pendente' : 'Pago'
        };
    }

    /**
     * Calcula número de diárias
     * @param {string} checkin - Data de check-in
     * @param {string} checkout - Data de check-out
     * @returns {number} Número de diárias
     */
    calculateNights(checkin, checkout) {
        if (!checkin || !checkout) return 0;
        
        const checkinDate = new Date(checkin);
        const checkoutDate = new Date(checkout);
        const diffTime = Math.abs(checkoutDate - checkinDate);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        
        return diffDays > 0 ? diffDays : 0;
    }

    /**
     * Formata data para exibição
     * @param {string} dateString - Data em formato ISO
     * @returns {string} Data formatada
     */
    formatDate(dateString) {
        const date = new Date(dateString);
        return date.toLocaleDateString('pt-BR', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
    }

    /**
     * Formata valor monetário
     * @param {number} value - Valor numérico
     * @returns {string} Valor formatado em Real
     */
    formatCurrency(value) {
        return new Intl.NumberFormat('pt-BR', {
            style: 'currency',
            currency: 'BRL'
        }).format(value);
    }

    /**
     * Gera ID único para o voucher
     * @returns {string} ID do voucher
     */
    generateVoucherId() {
        const timestamp = Date.now();
        const random = Math.random().toString(36).substr(2, 5);
        return `VOUCHER-${timestamp}-${random}`.toUpperCase();
    }

    /**
     * Valida datas do formulário
     * @returns {boolean} True se válido
     */
    validateDates() {
        const checkinDate = document.getElementById('checkinDate');
        const checkoutDate = document.getElementById('checkoutDate');
        
        if (checkinDate.value && checkoutDate.value) {
            const checkin = new Date(checkinDate.value);
            const checkout = new Date(checkoutDate.value);
            
            if (checkout <= checkin) {
                checkoutDate.setCustomValidity('A data de check-out deve ser posterior à data de check-in');
                checkoutDate.focus();
                return false;
            } else {
                checkoutDate.setCustomValidity('');
                return true;
            }
        }
        return true;
    }

    /**
     * Gera o arquivo PDF do voucher
     * @param {Object} data - Dados do voucher
     */
    async generateVoucher(data) {
        try {
            // Gera PDF usando o gerador específico
            if (typeof PDFVoucherGenerator !== 'undefined') {
                const pdfGenerator = new PDFVoucherGenerator();
                await pdfGenerator.generatePDF(data);
            } else {
                // Fallback para Word/HTML se PDF não estiver disponível
                await this.generateFallbackDocument(data);
            }

        } catch (error) {
            console.error('Erro ao gerar documento:', error);
            // Fallback: gera documento HTML
            await this.generateHtmlDocument(null, data);
        }
    }

    /**
     * Gera documento de fallback (Word ou HTML)
     * @param {Object} data - Dados do voucher
     */
    async generateFallbackDocument(data) {
        try {
            // Tenta carregar o template Word primeiro
            const templateBuffer = await this.loadTemplate();
            
            // Verifica se é um template Word ou HTML
            if (this.isWordTemplate(templateBuffer)) {
                await this.generateWordDocument(templateBuffer, data);
            } else {
                await this.generateHtmlDocument(templateBuffer, data);
            }

        } catch (error) {
            console.error('Erro ao gerar documento de fallback:', error);
            // Último recurso: gera documento HTML
            await this.generateHtmlDocument(null, data);
        }
    }

    /**
     * Verifica se o template é um arquivo Word
     * @param {ArrayBuffer} buffer - Buffer do template
     * @returns {boolean} True se for Word
     */
    isWordTemplate(buffer) {
        // Verifica se contém a assinatura de arquivo Word (.docx)
        const uint8Array = new Uint8Array(buffer);
        const signature = String.fromCharCode(...uint8Array.slice(0, 4));
        return signature === 'PK\x03\x04'; // Assinatura ZIP (Word é um ZIP)
    }

    /**
     * Gera documento Word usando Docxtemplater
     * @param {ArrayBuffer} templateBuffer - Buffer do template Word
     * @param {Object} data - Dados do voucher
     */
    async generateWordDocument(templateBuffer, data) {
        if (typeof PizZip === 'undefined' || typeof Docxtemplater === 'undefined') {
            throw new Error('Bibliotecas Word não carregadas');
        }

        const zip = new PizZip(templateBuffer);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        // Substitui os placeholders pelos dados
        doc.setData(data);
        
        // Renderiza o documento
        doc.render();

        // Gera o arquivo final
        const buffer = doc.getZip().generate({
            type: 'blob',
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        // Faz o download
        this.downloadFile(buffer, `Voucher_${data.nome.replace(/\s+/g, '_')}_${data.voucherId}.docx`);
    }

    /**
     * Gera documento HTML
     * @param {ArrayBuffer|null} templateBuffer - Buffer do template HTML (opcional)
     * @param {Object} data - Dados do voucher
     */
    async generateHtmlDocument(templateBuffer, data) {
        let htmlContent;
        
        if (templateBuffer) {
            // Usa template HTML carregado
            const decoder = new TextDecoder();
            htmlContent = decoder.decode(templateBuffer);
        } else {
            // Usa template HTML inline
            htmlContent = await this.createInlineHtmlTemplate();
        }

        // Substitui placeholders pelos dados
        htmlContent = this.replacePlaceholders(htmlContent, data);

        // Converte HTML para Blob
        const blob = new Blob([htmlContent], { type: 'text/html' });
        
        // Faz o download
        this.downloadFile(blob, `Voucher_${data.nome.replace(/\s+/g, '_')}_${data.voucherId}.html`);
    }

    /**
     * Cria template HTML inline
     * @returns {string} Template HTML
     */
    async createInlineHtmlTemplate() {
        const templateBuffer = await this.loadHtmlTemplate();
        const decoder = new TextDecoder();
        return decoder.decode(templateBuffer);
    }

    /**
     * Substitui placeholders no template
     * @param {string} template - Template com placeholders
     * @param {Object} data - Dados para substituição
     * @returns {string} Template com dados substituídos
     */
    replacePlaceholders(template, data) {
        let result = template;
        
        // Substitui todos os placeholders
        Object.keys(data).forEach(key => {
            const placeholder = `{{${key}}}`;
            const value = data[key] || '';
            result = result.replace(new RegExp(placeholder, 'g'), value);
        });

        // Substitui placeholders de status com classes CSS
        result = result.replace(/status-badge \$\{this\.getStatusClass\('{{status}}'\)\}/g, 
            `status-badge ${data.status === 'Pago' ? 'status-paid' : 'status-pending'}`);

        return result;
    }

    /**
     * Carrega o template do voucher
     * @returns {ArrayBuffer} Buffer do template
     */
    async loadTemplate() {
        try {
            const response = await fetch(this.templateUrl);
            if (!response.ok) {
                throw new Error(`Erro ao carregar template: ${response.status}`);
            }
            return await response.arrayBuffer();
        } catch (error) {
            console.error('Erro ao carregar template:', error);
            // Se não conseguir carregar o template, cria um template básico
            return this.createBasicTemplate();
        }
    }

    /**
     * Cria um template básico se o arquivo não estiver disponível
     * @returns {ArrayBuffer} Template básico
     */
    createBasicTemplate() {
        // Carrega o template HTML como fallback
        return this.loadHtmlTemplate();
    }

    /**
     * Carrega o template HTML como fallback
     * @returns {Promise<ArrayBuffer>} Template HTML
     */
    async loadHtmlTemplate() {
        try {
            // Tenta carregar o template específico do hotel primeiro
            let response = await fetch('../assets/templates/voucher_template_hotel.html');
            if (response.ok) {
                return await response.arrayBuffer();
            }
            
            // Fallback para o template genérico
            response = await fetch('../assets/templates/voucher_template_hotel.html');
            if (response.ok) {
                return await response.arrayBuffer();
            }
        } catch (error) {
            console.warn('Template HTML não encontrado, usando template inline');
        }
        
        // Template HTML básico inline como último recurso
        const htmlTemplate = `
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
                <meta charset="UTF-8">
                <title>Voucher de Hospedagem</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 40px; background: #f8f9fa; }
                    .container { max-width: 800px; margin: 0 auto; background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
                    .header { text-align: center; margin-bottom: 30px; background: linear-gradient(135deg, #3B82F6 0%, #8B5CF6 100%); color: white; padding: 30px; border-radius: 8px; }
                    .title { font-size: 28px; font-weight: bold; margin: 0; }
                    .subtitle { font-size: 16px; opacity: 0.9; margin-top: 10px; }
                    .voucher-id { background: #f3f4f6; padding: 15px; border-radius: 8px; text-align: center; font-weight: bold; color: #3B82F6; margin-bottom: 30px; }
                    .section { margin-bottom: 30px; }
                    .section-title { font-size: 20px; font-weight: bold; color: #1f2937; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 2px solid #e5e7eb; }
                    .field-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 20px; }
                    .field { display: flex; flex-direction: column; }
                    .field-label { font-weight: bold; color: #6b7280; font-size: 14px; margin-bottom: 5px; }
                    .field-value { font-size: 16px; color: #1f2937; padding: 10px; background: #f9fafb; border-radius: 6px; border-left: 4px solid #3B82F6; }
                    .values-section { background: #f8fafc; padding: 25px; border-radius: 8px; border: 1px solid #e2e8f0; }
                    .value-item { display: flex; justify-content: space-between; align-items: center; padding: 12px 0; border-bottom: 1px solid #e2e8f0; }
                    .value-item:last-child { border-bottom: none; }
                    .value-label { font-weight: 600; color: #374151; }
                    .value-amount { font-weight: bold; font-size: 18px; }
                    .value-total { color: #3B82F6; }
                    .value-remaining { color: #10B981; }
                    .status-badge { display: inline-block; padding: 8px 16px; border-radius: 20px; font-weight: bold; font-size: 14px; }
                    .status-paid { background: #d1fae5; color: #065f46; }
                    .status-pending { background: #fef3c7; color: #92400e; }
                    .observations { background: #fefce8; padding: 20px; border-radius: 8px; border-left: 4px solid #f59e0b; }
                    .footer { background: #f3f4f6; padding: 25px; text-align: center; color: #6b7280; font-size: 14px; margin-top: 30px; border-radius: 8px; }
                    .company { font-weight: bold; color: #3B82F6; margin-bottom: 5px; }
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <div class="title">VOUCHER DE HOSPEDAGEM</div>
                        <div class="subtitle">VoucherPro - Sistema de Gestão</div>
                    </div>
                    
                    <div class="voucher-id">ID do Voucher: {{voucherId}}</div>
                    
                    <div class="section">
                        <div class="section-title">Informações do Hóspede</div>
                        <div class="field-grid">
                            <div class="field">
                                <div class="field-label">Nome Completo</div>
                                <div class="field-value">{{nome}}</div>
                            </div>
                            <div class="field">
                                <div class="field-label">Tipo de Acomodação</div>
                                <div class="field-value">{{tipoAcomodacao}}</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="section">
                        <div class="section-title">Datas de Hospedagem</div>
                        <div class="field-grid">
                            <div class="field">
                                <div class="field-label">Data de Check-in</div>
                                <div class="field-value">{{checkin}}</div>
                            </div>
                            <div class="field">
                                <div class="field-label">Data de Check-out</div>
                                <div class="field-value">{{checkout}}</div>
                            </div>
                            <div class="field">
                                <div class="field-label">Quantidade de Diárias</div>
                                <div class="field-value">{{quantidadeDiarias}} diária(s)</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="section">
                        <div class="section-title">Valores</div>
                        <div class="values-section">
                            <div class="value-item">
                                <span class="value-label">Valor da Diária</span>
                                <span class="value-amount">{{valorDiaria}}</span>
                            </div>
                            <div class="value-item">
                                <span class="value-label">Valor Antecipado</span>
                                <span class="value-amount">{{valorAntecipado}}</span>
                            </div>
                            <div class="value-item">
                                <span class="value-label">Valor Total</span>
                                <span class="value-amount value-total">{{total}}</span>
                            </div>
                            <div class="value-item">
                                <span class="value-label">Valor Restante</span>
                                <span class="value-amount value-remaining">{{restante}}</span>
                            </div>
                            <div class="value-item">
                                <span class="value-label">Status</span>
                                <span class="status-badge ${this.getStatusClass('{{status}}')}">{{status}}</span>
                            </div>
                        </div>
                    </div>
                    
                    <div class="section">
                        <div class="section-title">Observações</div>
                        <div class="observations">{{observacoes}}</div>
                    </div>
                    
                    <div class="footer">
                        <div class="company">VoucherPro</div>
                        <div>Voucher gerado em {{dataEmissao}}</div>
                        <div>Sistema de Gestão de Vouchers de Hospedagem</div>
                    </div>
                </div>
            </body>
            </html>
        `;

        const encoder = new TextEncoder();
        return encoder.encode(htmlTemplate).buffer;
    }

    /**
     * Retorna a classe CSS para o status
     * @param {string} status - Status do voucher
     * @returns {string} Classe CSS
     */
    getStatusClass(status) {
        return status === 'Pago' ? 'status-paid' : 'status-pending';
    }

    /**
     * Faz o download do arquivo
     * @param {Blob} blob - Arquivo para download
     * @param {string} filename - Nome do arquivo
     */
    downloadFile(blob, filename) {
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
    }

    /**
     * Mostra estado de carregamento
     */
    showLoadingState() {
        const confirmButton = document.getElementById('confirmButton');
        if (confirmButton) {
            confirmButton.innerHTML = '<i class="fas fa-spinner fa-spin mr-3"></i>Gerando PDF...';
            confirmButton.disabled = true;
        }
    }

    /**
     * Esconde estado de carregamento
     */
    hideLoadingState() {
        const confirmButton = document.getElementById('confirmButton');
        if (confirmButton) {
            confirmButton.innerHTML = '<i class="fas fa-check mr-3"></i>Confirmar Voucher';
            confirmButton.disabled = false;
        }
    }

    /**
     * Mostra modal de sucesso
     */
    showSuccessModal() {
        const successModal = document.getElementById('successModal');
        if (successModal) {
            successModal.classList.remove('hidden');
            successModal.classList.add('flex');
        }
    }

    /**
     * Mostra mensagem de erro
     * @param {string} message - Mensagem de erro
     */
    showError(message) {
        // Cria notificação de erro
        const notification = document.createElement('div');
        notification.className = 'fixed top-4 right-4 p-4 bg-red-500 text-white rounded-xl shadow-lg z-50 animate-slide-up';
        notification.innerHTML = `
            <div class="flex items-center space-x-2">
                <i class="fas fa-exclamation-triangle"></i>
                <span>${message}</span>
            </div>
        `;
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            notification.style.animation = 'slideUp 0.3s ease-out reverse';
            setTimeout(() => {
                notification.remove();
            }, 300);
        }, 5000);
    }
}

// Inicializa o gerador quando o DOM estiver carregado
document.addEventListener('DOMContentLoaded', () => {
    // Verifica se as bibliotecas necessárias estão carregadas
    if (typeof PizZip === 'undefined' || typeof Docxtemplater === 'undefined') {
        console.warn('Bibliotecas PizZip ou Docxtemplater não encontradas. Carregando fallback...');
        // Inicializa com funcionalidade limitada
        window.voucherGenerator = new VoucherGenerator();
    } else {
        // Inicializa com funcionalidade completa
        window.voucherGenerator = new VoucherGenerator();
    }
});

// Exporta a classe para uso global
window.VoucherGenerator = VoucherGenerator;
