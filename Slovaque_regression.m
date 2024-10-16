% pr�vision de charge avec la m�thode de la regression lin�aire multiple
%  R�seau Slovaque 1998-2011
close all;
clear all;
state0 = 0;
rand('state',state0);
donnees = xlsread('load slovakia.xls',2,'A13:B26'); % charger les done�s des charges (GWh)
w=round(0.6*length(donnees));
charge=donnees(:,2);
charge_precedente=donnees(1:w,2);
charge_actuelle=donnees(2:w+1,2);
x=[charge_precedente];
 y=charge_actuelle;
 c=ones(size(x(:,1)));
  X=[c x];
  b = REGRESS(y,X)
  b0=b(1)
%   R=CORRCOEF(x,y)
 yestime=b'*X';
 erreur=y-yestime';
set(figure,'Color','white')
 plot(yestime,y,'+'); grid
 hold on
 plot(yestime,yestime,'r');
 x_labels = get(gca, 'XTick');
 set(gca, 'XTickLabel', x_labels);
 y_labels = get(gca, 'YTick');
 set(gca, 'YTickLabel', y_labels);
 title('Corr�lation entre la charge dis�r�e et r�alis�e ','FontSize',11, 'FontName','Times New Roman','FontWeight','normal')
 xlabel(' Charge Pr�vue (GWh)','FontSize',11, 'FontName','Times New Roman','FontWeight','normal')
 ylabel(' Charge R�alis�e (GWh)','FontSize',11, 'FontName','Times New Roman','FontWeight','normal')
 legend('Pr�vision','Tendance')
%  Phase de pr�vision
charge_precedente=donnees(w+1:end-1,2);
charge_actuelle=donnees(w+2:end,2);
x=[charge_precedente];
y=charge_actuelle;
c=ones(size(x(:,1)));
X=[c x];
yprev=b'*X';
[yprev' y];
re=yprev';
de=y;
yprv=yprev';
year=donnees(w+2:end,1);
disp ('ACTUAL AND FORECAST HOURLY LOAD USING LINEAR REGRESSION');
    disp('            ');
    fprintf('\nYear         Actual    Forecast      ');
    fprintf('\nYear           GWh       GWh         ');
disp('            ');
disp('            ');
for i=1:length(year)
        fprintf('%-08.2d %-10.5g %-7.5g\n',year(i),y(i),yprv(i))
        m(i,:)=[year(i) y(i) yprv(i)];
end
% Erreur moyenne absolue de pr�vision
ct=sum(abs((de-re)./de));
   MAPE=ct.*100/length(re);
   fprintf('\n MAPE = ');fprintf('%-2.2f',MAPE); fprintf(' %% \n')
    mape={'MAPE(%)'};
set(figure,'Color','white')
plot(datenum(year,1,1),de,'b')
hold on
plot(datenum(year,1,1),re,'K')
datetick('x','yyyy') 
y_labels = get(gca, 'YTick');
set(gca, 'YTickLabel', y_labels);
title('Pr�vision de charge par la r�gression ','FontSize',11, 'FontName','Times New Roman','FontWeight','normal')
xlabel('Ann�e','FontSize',11, 'FontName','Times New Roman','FontWeight','normal')
ylabel('Charge (GWh)','FontSize',11, 'FontName','Times New Roman','FontWeight','normal')
legend('Dis�r�e ','Pr�vue')
legend('boxoff')
d = {'Ann�e' 'Actuelle(GWh)'  'Pr�vue(GWh)'};
z= {'Pr�vision de charge par la r�gression'};
xlswrite('Pr�vision de charge R�seau Slovaque',z,'LINEAR REGRESSION','C1')
xlswrite('Pr�vision de charge R�seau Slovaque',d, 'LINEAR REGRESSION','C2')
xlswrite('Pr�vision de charge R�seau Slovaque',m, 'LINEAR REGRESSION','C3') 
xlswrite('Pr�vision de charge R�seau Slovaque',mape, 'LINEAR REGRESSION','K3')
xlswrite('Pr�vision de charge R�seau Slovaque',MAPE, 'LINEAR REGRESSION','K4')