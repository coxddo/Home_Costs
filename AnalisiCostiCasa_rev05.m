%% Script per il calcolo delle spese di mantenimento annuali di un appartamento, dati EPGL e i vari costi
%% rev02 - messo costo energia medio, tiene conto anche del costi vari e iva
%% rev03 - agginto un plot mensile dei costi per materia
%% rev04 - aggiunto input da excel
%% rev05 - aggiunto input da excel delle tipologie di elettrodomestici

tic
delete excelInput.xlsx
delete excelOutput.xlsx
delete AnalisiCostiCasa.pptx
% rmdir Immagini
clc, close all, clear

mkdir Immagini
a=dir;
%% Input - PrezziMateria
prezziMateriaInput = readtable('INPUT.xlsx', 'Sheet','PrezziMateria');
prezzoElettricita = prezziMateriaInput.Elettricit_; % €/kWh - monooraria + iva + spese trasporto - esagerato
% costoCanoneReteElettricita = 34.2; % €/anno
% costoSpeseCommercializzazione = 16.2; % €/anno
prezzoGasMetano = prezziMateriaInput.Metano; % €/Sm3
prezzoInternetMese = prezziMateriaInput.Internet; % €/mese
%% Input - CASA
prezziCasaInput = readtable('INPUT.xlsx', 'Sheet','PrezziCasa');

EPGLnren = prezziCasaInput.EPGL; % kWh/m2/anno
superficie = prezziCasaInput.Area; % m2
prezzoAffittoMese = prezziCasaInput.Affitto; % €/mese
prezzoCondominioMese = prezziCasaInput.Condominio; % €/mese
% rendimentoCaldaia = .8; % -
% Hi_metano = 9.5833; % kWh/Smc
%% INPUT - Quantità Elettrodomestici
elettrodomestici = readtable('INPUT.xlsx', 'Sheet','Elettrodomestici');
nForno = elettrodomestici.nForno;
nCellulari = elettrodomestici.nCellulari;
nLavatrici = elettrodomestici.nLavatrici;
nTelevisori = elettrodomestici.nTelevisori;
nFrigo = elettrodomestici.nFrigo;
nScaldabagno = elettrodomestici.nScaldabagno; %%%
nPcPortatile = elettrodomestici.nPcPortatile;

%% Costanti
HiMetano = 11; % kWh/Sm3  (calore specifico del metano)

%% Creazione tabella Input
% Definizione dei nomi delle colonne
colonne = {'Descrizione', 'Valore', 'UM'};

% Creazione della struttura dati per la tabella
input_data = struct(...
    'Descrizione', { ...
        'Prezzo Energia Elettrica', ...
        'Prezzo Gas Metano', ...
        'Prezzo Internet Mensile', ...
        'EPGLnren', ...
        'Superficie Appartamento', ...
        'Prezzo Affitto Mensile', ...
        'Prezzo Condominio Mensile' ...
    }, ...
    'Valore', { ...
        prezzoElettricita, ...
        prezzoGasMetano, ...
        prezzoInternetMese, ...
        EPGLnren, ...
        superficie, ...
        prezzoAffittoMese, ...
        prezzoCondominioMese ...
    }, ...
    'UM', { ...
        '€/kWh', ...
        '€/Sm3', ...
        '€/mese', ...
        'kWh/m2/anno', ...
        'm2', ...
        '€/mese', ...
        '€/mese' ...
    } ...
);

% Conversione della struttura dati in tabella
T = struct2table(input_data);
% T.Valore = round(T.Valore, 2); % massimo 2 cifre decimali

%% Stampa della tabella
disp('Tabella Input:')
disp(T)  % Stampa la tabella a console
writetable(T, 'excelInput.xlsx')
% COSTI 
%% En. elettrica

utenze.forno = nForno;
utenze.cellulare = nCellulari;
utenze.lavatrice = nLavatrici;
utenze.televisore = nTelevisori;
utenze.frigo = nFrigo;
utenze.scaldabagno = nScaldabagno;
utenze.pcportatile = nPcPortatile;

[potenze, consumi, tempi, costoElettricitaAnno] = costoEnergiaElettrica(nForno, nCellulari, nLavatrici, nTelevisori, nFrigo, nScaldabagno, nPcPortatile, prezzoElettricita);
costoTotElettricitaAnno= costoElettricitaAnno; % + costoSpeseCommercializzazione + costoCanoneReteElettricita;

%% En. termica
[consumoEnMetano, consumoVolMetano, costoMetanoAnno] = costoGasMetano(superficie, EPGLnren, prezzoGasMetano);

%% Internet
costoInternetAnno = Mensile2Anno(prezzoInternetMese);

%% Condominio
costoCondominioAnno = Mensile2Anno(prezzoCondominioMese);

%% Affitto
CostoAffittoAnnuo   = Mensile2Anno(prezzoAffittoMese);

%% Calcolo COSTI da CONSUMI - aka costi di mantenimento annuo
ConsumiTotali.elettricita = sum(struct2array(consumi).*prezzoElettricita);
ConsumiTotali.metano = costoMetanoAnno;
ConsumiTotali.telefono = costoInternetAnno;
ConsumiTotali.condominio = costoCondominioAnno;
CostoMantenimentoAnnuo = sum(struct2array(ConsumiTotali));

%% Calcolo COSTI totali 
[costi, CostoComplessivoAnnuo] = SpeseComplessiveCasa(CostoMantenimentoAnnuo, CostoAffittoAnnuo, 1);

%% PLOT
fig5 = figure("Name", "Costi totali annui");
    subplot(1,2,1)        
    bar("Costi annui", struct2array(costi), 'stacked');
    barColors = get(gca,'ColorOrder'); % Get colors from bar chart
    legend(fieldnames(costi), 'Location','eastoutside');
    
    subplot(1,2,2)
    piechart(struct2array(costi), 'Color', barColors); % Apply the same colors to pie chart
    sgtitle("Costi totali annui")

fig6 = figure("Name", "Costi totali mensili");
    subplot(1,2,1)        
    bar("Costi annui", struct2array(costi)/12, 'stacked');
    barColors = get(gca,'ColorOrder'); % Get colors from bar chart
    legend(fieldnames(costi), 'Location','eastoutside');
    
    subplot(1,2,2)
    piechart(struct2array(costi)/12, 'Color', barColors); % Apply the same colors to pie chart
    sgtitle("Costi totali mensili")
    

fig1 = figure("Name", "Consumi elettrici e tipi di utenze");
    subplot(1,5, 1)
        bar(["Utenze"], [struct2array(utenze)], 'stacked');
        legend(fieldnames(utenze), 'Location','westoutside');
    subplot(1,5, 2)
        bar(["Potenze installate"], [struct2array(potenze)/1000], 'stacked');
    subplot(1,5, 3)
        bar(["Tempi"], [struct2array(tempi)], 'stacked');
    subplot(1,5, 4)
        bar(["Consumi"], [struct2array(consumi)], 'stacked');
    subplot(1,5, 5)
        bar(["Costi ANNO"], struct2array(consumi).*prezzoElettricita, 'stacked');
        sgtitle("Costi Elettricità ANNO")

fig2 = figure("Name", "Consumo Metano");
    subplot(1,2,1)
        bar(["Consumi Metano"], [consumoVolMetano], 'stacked');
            ylabel("St. metri cubi");
    subplot(1,2,2)
        bar(["Costi Metano"], [costoMetanoAnno], 'stacked');
            ylabel("€ metano/anno");
          %%
fig3 =  figure("Name","Costi x Materia Annui");
     % yyaxis left
    k=bar("Consumi", [struct2array(ConsumiTotali)], 'stacked');
    text(k(1).XEndPoints, k(end).YEndPoints+250, [num2str(round(k(end).YEndPoints,0)), ' €'], 'FontWeight', 'bold', 'FontSize', 20);
    hold on
    z=bar("Affitto", CostoAffittoAnnuo);
    text(z(1).XEndPoints, z(end).YEndPoints+250, [num2str(round(z(end).YEndPoints,0)), ' €'], 'FontWeight', 'bold', 'FontSize', 20);
    % hold on
    % plot(sum(struct2array(ConsumiTotali)), '*')
    title('Costi ANNUI x materia');
    grid on
    legend([fieldnames(ConsumiTotali)], 'Location', 'northwest')
    ylim([0 z(end).YEndPoints+500])
    ylabel ('€/anno')
    
%%
    fig7 =  figure("Name","Costi MESE x Materia ");
     % yyaxis left
    k=bar("Consumi", [struct2array(ConsumiTotali)/12], 'stacked');
    text(k(1).XEndPoints, k(end).YEndPoints+50, [num2str(round(k(end).YEndPoints,0)), ' €'], 'FontWeight', 'bold', 'FontSize', 20);
    hold on
    z=bar("Affitto", CostoAffittoAnnuo/12);
    text(z(1).XEndPoints, z(end).YEndPoints+50, [num2str(round(z(end).YEndPoints,0)), ' €'], 'FontWeight', 'bold', 'FontSize', 20);
    % hold on
    % plot(sum(struct2array(ConsumiTotali)), '*')
    title('Costi MESE x materia');
    grid on
    legend([fieldnames(ConsumiTotali)], 'Location', 'northwest')
    ylim([0 z(end).YEndPoints+500])
    ylabel ('€/anno')
    
    %%
fig4 = figure("Name","Torta dei costi");
    pie(struct2array(ConsumiTotali));
    legend(fieldnames(ConsumiTotali), 'Location','eastoutside')


% stampa come svg
PrintAsImage(fig1, fig2, fig3, fig4, fig5, fig6, fig7, a)



%% Stampa Dati Output
% Create a structure for the output data
output_data = struct(...
    'Descrizione', { ...
        'Costo Elettricita Anno', ...
        'Costo Metano Anno', ...
        'Costo Internet Anno', ...
        'Costo Condominio Anno', ...
        'Costo Affitto Annuo', ...
        'Costo Manutenzione Anno', ...
        'Costo Complessivo Anno' ...
    }, ...
    'Valore', { ...
        costoTotElettricitaAnno, ...
        costoMetanoAnno, ...
        costoInternetAnno, ...
        costoCondominioAnno, ...
        CostoAffittoAnnuo, ...
        CostoMantenimentoAnnuo, ...
        CostoComplessivoAnnuo ...
    }, ...
    'UM', { ...
        '€/anno', ...
        '€/anno', ...
        '€/anno', ...
        '€/anno', ...
        '€/anno', ...
        '€/anno', ...
        '€/anno' ...
    } ...
);

% Convert the structure to a table
T_output = struct2table(output_data);
% T_output.Valore = round(T_output.Valore, 2); % massimo 2 cifre decimali

% Display the output table
disp('Tabella Output:')
disp(T_output)  % Print the table to the console
writetable(T_output, 'excelOutput.xlsx')
close all

% stampa nel ppt
run PPTcreator.m
toc

%% Funzioni esterne

function [potenze, consumi, tempi, costo] = costoEnergiaElettrica(nForno, nCellulari, nLavatrici, nTelevisori, nFrigo, nScaldabagno, nPcPortatile, prezzoKWh)
    % Calcola il costo dell'energia elettrica per l'utilizzo di diversi elettrodomestici.
    % Potenza media di ogni elettrodomestico
    potenzaForno = 2000; % Watt
    potenzaCellulare = 5; % Watt
    potenzaLavatrice = 2000; % Watt
    potenzaTelevisore = 100; % Watt
    potenzaFrigo = 100; % Watt/ 1000
    potenzaScaldabagno = 2000; % Watt
    potenzaPcPortatile = 65; % Watt
    
    potenze.forno = potenzaForno;
    potenze.cellulare = potenzaCellulare;
    potenze.lavatrice = potenzaLavatrice;
    potenze.televisore = potenzaTelevisore;
    potenze.frigo = potenzaFrigo;
    potenze.scaldabagno = potenzaScaldabagno;
    potenze.pcportatile = potenzaPcPortatile;
    
    % Tempo di utilizzo stimato al giorno
    tempoForno = 2/7; % ora/giorno
    tempoCellulare = 3; % ore/giorno
    tempoLavatrice =  4/7; % ore/giorno
    tempoTelevisore = 1; % ore/giorno
    tempoFrigo = 8; % ore/giorno
    tempoScaldabagno = 1; % ora/giorno
    tempoPcPortatile = 12; % ore/giorno
    
    tempi.forno = tempoForno;
    tempi.cellulare = tempoCellulare;
    tempi.lavatrice = tempoLavatrice;
    tempi.televisore = tempoTelevisore;
    tempi.frigo = tempoFrigo;
    tempi.scaldabagno = tempoScaldabagno;
    tempi.pcportatile = tempoPcPortatile;
    % Calcolo del consumo energetico per ogni elettrodomestico per anno
    consumoForno = nForno * potenzaForno * tempoForno * 365 / 1000; % kWh
    consumoCellulare = nCellulari * potenzaCellulare * tempoCellulare * 365 / 1000; % kWh
    consumoLavatrice = nLavatrici * potenzaLavatrice * tempoLavatrice * 365 / 1000; % kWh
    consumoTelevisore = nTelevisori * potenzaTelevisore * tempoTelevisore * 365 / 1000; % kWh
    consumoFrigo = nFrigo * potenzaFrigo * tempoFrigo * 365 / 1000; % kWh
    consumoScaldabagno = nScaldabagno * potenzaScaldabagno * tempoScaldabagno * 365 / 1000; % kWh
    consumoPcPortatile = nPcPortatile * potenzaPcPortatile * tempoPcPortatile * 365 / 1000; % kWh
    
    consumi.forno = consumoForno;
    consumi.cellulare = consumoCellulare;
    consumi.lavatrice = consumoLavatrice;
    consumi.televisore = consumoTelevisore;
    consumi.frigo = consumoFrigo;
    consumi.scaldabagno = consumoScaldabagno;
    consumi.pcportatile = consumoPcPortatile;

    % Consumo energetico totale
    consumoTotale = consumoForno + consumoCellulare + consumoLavatrice + consumoTelevisore + consumoFrigo + consumoScaldabagno + consumoPcPortatile; % kWh
    % Prezzo dell'energia elettrica
    % prezzoKWh = 0.20; % euro/kWh
    % Calcolo del costo dell'energia elettrica
    costo = consumoTotale * prezzoKWh;
    
end
function [energiaPrimariaAnno, volumeMetanoAnno, costo_gas] = costoGasMetano(superficie, EPGLnren, costo_smc)


% potere calorifico metano 
HiMetano = 11; % kWh/m3

% Calcolo del consumo annuale di gas metano
energiaPrimariaAnno = superficie * EPGLnren; % / rendimentoCaldaia;

% Conversione da mc a smc
volumeMetanoAnno = energiaPrimariaAnno / HiMetano;

% Calcolo del costo annuale del gas metano
costo_gas = volumeMetanoAnno * costo_smc;

end

function costoAnnuo = Mensile2Anno(prezzoMensile)
       costoAnnuo = prezzoMensile*12;
end

function [costi, CostoComplessivoAnnuo] = SpeseComplessiveCasa(CostoMantenimentoAnnuo, CostoAffittoAnnuo, print)
    CostoComplessivoAnnuo = CostoMantenimentoAnnuo + CostoAffittoAnnuo;
    costi.mantenimento = CostoMantenimentoAnnuo;
    costi.affitto = CostoAffittoAnnuo;    
end

function PrintAsImage(fig1, fig2, fig3, fig4, fig5, fig6, fig7, a)
    directory= a.folder;
    directory = [directory, '\Immagini'];
    l=cd(directory);
    saveas(fig1, 'ConsElettrici.svg')
    saveas(fig2, 'ConsMetano.svg')
    saveas(fig3, 'CostiConsumi.svg')
    saveas(fig4, 'TortaDeiCosti.svg')
    saveas(fig5, 'CostiTotaliAnnui.svg')
    saveas(fig6, 'CostiTotaliMensili.svg')
    saveas(fig7, 'CostiMESEMateria.svg')
    cd(l)
end
