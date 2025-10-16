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
    <title>Voucher Hotel Porto da Lua</title>
    <style>
        * {
            box-sizing: border-box;
        }

        body {
            margin: 16px;
            background: #fff;
            color: #000;
            font-family: "Times New Roman", Times, serif;
            font-size: 14px;
            line-height: 1.2;
        }

        .voucher-container {
            width: 746px;
            border: 1px solid #000;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            margin-top: 5rem;
        }

        .voucher-row {
            display: flex;
            width: 100%;
            border-bottom: 1px solid #000;
        }

        .voucher-cell {
            border-right: 1px solid #000;
            padding: 8px;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        .voucher-cell:last-child {
            border-right: none;
        }

        /* Header */
        .header {
            display: flex;
            width: 100%;
            height: 106px;
            border-bottom: 1px solid #000;
        }

        .header-logo {
            width: 146px;
            text-align: center;
            position: relative;
            border-right: 1px solid #000;
        }

        .header-logo-placeholder {
            width: 100px;
            height: 90px;
            margin: 8px auto 0;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 12px;
        }

        .header-logo img {
            position: absolute;
            top: 8px;
            left: 50%;
            transform: translateX(-50%);
            width: 100px;
            height: 90px;
            object-fit: contain;
            opacity: 1;
        }

        .header-info {
            flex: 1;
            text-align: center;
            padding: 8px 9px 0 9px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        .hotel-name {
            font-weight: 800;
            text-transform: uppercase;
            font-size: 20px;
            margin: 0;
        }

        .hotel-address {
            font-size: 14px;
            margin-top: 6px;
        }

        /* Título do voucher */
        .voucher-title {
            height: 50px;
            text-align: center;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.2px;
            padding-top: 6px;
            border-bottom: 1px solid #000;
        }

        /* Check-in / Check-out */
        .main-info {
            display: flex;
            width: 100%;
            height: 60px;
            border-bottom: 1px solid #000;
        }

        .main-info .label-cell {
            width: 146px;
            font-weight: 700;
            text-transform: uppercase;
            display: flex;
            align-items: center;
            justify-content: center;
            border-right: 1px solid #000;
        }

        .main-info .label-cell1 {
            width: 130px;
            font-weight: 700;
            text-transform: uppercase;
            display: flex;
            align-items: center;
            justify-content: center;
            border-right: 1px solid #000;
        }

        .main-info .info-cell {
            flex: 1;
            display: flex;
            padding: 0;
        }

        .split-left {
            border-right: 1px solid #000;
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 9px;
        }

        .split-right {
            flex: 2;
            display: grid;
            grid-template-columns: 250px 1fr;
        }

        .col-2-label {
            font-weight: 700;
            text-transform: uppercase;
            display: flex;
            align-items: center;
            padding-left: 9px;
        }

        .value-cell {
            display: flex;
            align-items: center;
            padding-left: 12px;
            white-space: nowrap;
        }

        /* Hóspedes */
        .hospedes {
            display: flex;
            align-items: center;
            height: 4.5rem;
            padding: 4px 8px;
        }

        .hospedes .label-cell {
            font-weight: 700;
            text-transform: uppercase;
            margin-right: 12px;
        }

        /* Pagamento */
        .payment-title {
            height: 32px;
            text-align: center;
            font-weight: 800;
            text-transform: uppercase;
            padding-top: 6px;
            border-top: 1px solid #000;
            border-bottom: 1px solid #000;
        }

        .payment-section {
            display: flex;
            width: 100%;
        }

        .policy-cell {
            width: 50%;
            padding: 10px 14px 12px 14px;
        }

        .policy-cell strong {
            display: inline-block;
            margin-bottom: 6px;
        }

        .policy-cell ul {
            margin: 8px 0 0 0;
            padding-left: 18px;
        }

        .policy-cell li {
            margin: 7px 0;
        }

        .values-table {
            width: 50%;
            border-collapse: collapse;
        }

        .values-table-row {
            display: flex;
            border-bottom: 1px solid #000;
            border-left: 1px solid #000;
            height: 36px;
        }

        .val-label {
            flex: 2;
            font-weight: 700;
            text-transform: uppercase;
            display: flex;
            align-items: center;
            padding-left: 12px;
        }

        .val-currency {
            width: 28px;
            text-align: center;
            flex-shrink: 0;
        }

        .val-amount {
            flex: 1;
            text-align: right;
            padding-right: 12px;
            white-space: nowrap;
        }

        /* Rodapé */
        .footer-notes {
            padding: 10px 12px;
            font-weight: 800;
            line-height: 1.5;
            height: 84px;
            font-size: 16px;
            border-top: 1px solid #000;
        }

        /* Assinatura e data */
        .signature-date {
            display: flex;
            width: 100%;
            height: 68px;
            border-top: 1px solid #000;
        }

        .signature-box {
            flex: 2;
            position: relative;
            border-right: 1px solid #000;
        }

        .signature-hint {
            position: absolute;
            left: 10px;
            top: 10px;
            font-size: 12px;
            color: #333;
        }

        .assinatura-overlay {
            position: absolute;
            left: 10px;
            width: 250px;
            height: 70px;
            object-fit: contain;
            opacity: 1;
        }

        .date-cell {
            flex: 1;
            text-align: center;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 14px;
        }

        .excessao {
            border-bottom: none;
        }

        /* ===== BOTÃO DE GERAR PDF ===== */
        #gerarPdf {
            position: fixed;
            top: 20px;
            right: 20px;
            background-color: #007bff;
            color: #fff;
            border: none;
            padding: 10px 18px;
            border-radius: 8px;
            cursor: pointer;
            font-size: 15px;
            font-weight: bold;
            box-shadow: 0 3px 8px rgba(0, 0, 0, 0.2);
            transition: all 0.3s ease;
            z-index: 1000;
        }

        #gerarPdf:hover {
            background-color: #0056b3;
            transform: scale(1.05);
        }

        @media print {
            #gerarPdf {
                display: none;
            }
        }
    </style>
</head>

<body>
    <div class="voucher-container">
        <!-- Header -->
        <div class="header">
            <div class="header-logo">
                <div class="header-logo-placeholder"></div>
                <img src="../img/logo.png" alt="Logo Hotel Porto da Lua" class="logo-overlay">
            </div>
            <div class="header-info">
                <div class="hotel-name">HOTEL PORTO DA LUA - MARUDÁ</div>
                <div class="hotel-address">Av. Beira Mar – Alegre – Marudá</div>
                <div class="hotel-address">Marapanim – Pará – Brasil</div>
                <div class="hotel-address">Fone: (91) 9 9369-1074</div>
            </div>
        </div>

        <!-- Título do voucher -->
        <div class="voucher-title">VOUCHER- QUARTO X</div>

        <!-- Check-in -->
        <div class="main-info">
            <div class="label-cell1">CHECK IN</div>
            <div class="info-cell">
                <div class="split-left">{{checkin}}</div>
                <div class="split-right">
                    <div class="col-2-label">TIPO DE ACOMODAÇÃO:</div>
                    <div class="value-cell">{{tipoAcomodacao}}</div>
                </div>
            </div>
        </div>

        <!-- Check-out -->
        <div class="main-info">
            <div class="label-cell1">CHECK OUT</div>
            <div class="info-cell">
                <div class="split-left">{{checkout}}</div>
                <div class="split-right">
                    <div class="col-2-label">OBSERVAÇÃO:</div>
                    <div class="value-cell">{{observacoes}}</div>
                </div>
            </div>
        </div>

        <!-- Hóspedes -->
        <div class="hospedes">
            <div class="label-cell">HÓSPEDES:</div>
            <div class="value-cell">{{nome}}</div>
        </div>

        <!-- Pagamento -->
        <div class="payment-title">DADOS DO PAGAMENTO</div>
        <div class="payment-section">
            <div class="policy-cell">
                <strong>Política de Cancelamento</strong>
                <ul style="list-style-type: none;">
                    <li>- 14 Dias antes: Reembolso Integral</li>
                    <li>- 7 dias antes: Crédito em diárias</li>
                    <li>- 48h: No Show Integral sem reembolso ou crédito</li>
                </ul>
            </div>
            <div class="values-table">
                <div class="values-table-row">
                    <div class="val-label">VALOR DA DIÁRIA</div>
                    <div class="val-currency">R$</div>
                    <div class="val-amount">{{valorDiaria}}</div>
                </div>
                <div class="values-table-row">
                    <div class="val-label">QUANTIDADE DE DIÁRIAS</div>
                    <div class="val-currency"></div>
                    <div class="val-amount">{{quantidadeDiarias}}</div>
                </div>
                <div class="values-table-row">
                    <div class="val-label">VALOR TOTAL</div>
                    <div class="val-currency">R$</div>
                    <div class="val-amount">{{total}}</div>
                </div>
                <div class="values-table-row">
                    <div class="val-label">VALOR ANTECIPADO</div>
                    <div class="val-currency">R$</div>
                    <div class="val-amount">{{valorAntecipado}}</div>
                </div>
                <div class="values-table-row excessao">
                    <div class="val-label">VALOR A PAGAR</div>
                    <div class="val-currency">R$</div>
                    <div class="val-amount">{{restante}}</div>
                </div>
            </div>
        </div>

        <!-- Observações -->
        <div class="footer-notes">
            OBS.: DIÁRIA INICIA AS 12:00 DA DATA DE ENTRADA E ENCERRA AS 12:00 DA DATA DE SAÍDA.<br>
            O HÓSPEDE DEVERÁ QUITAR O DÉBITO NO <strong>CHECK IN</strong>.
        </div>

        <!-- Assinatura e data -->
        <div class="signature-date">
            <div class="signature-box">
                <div class="signature-hint"></div>
                <img src="../img/assinatura.jfif" alt="Assinatura" class="assinatura-overlay">
            </div>
            <div class="date-cell">{{dataEmissao}}</div>
        </div>
    </div>
    <button id="gerarPdf" style="margin-bottom: 16px; padding: 8px 14px; font-size: 14px; cursor: pointer;">
        Gerar PDF
    </button>
    <!-- Biblioteca html2pdf -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <script>
        document.getElementById('gerarPdf').addEventListener('click', () => {
            const voucher = document.querySelector('.voucher-container');

            const opcoes = {
                margin: 0,
                filename: 'voucher-hotel.pdf',
                image: { type: 'jpeg', quality: 0.98 },
                html2canvas: { scale: 3, useCORS: true },
                jsPDF: { unit: 'pt', format: 'a4', orientation: 'portrait' }
            };

            html2pdf().set(opcoes).from(voucher).save();
        });
    </script>

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
