#!/usr/bin/perl
use strict;

use CGI;
use DBI;
use HTML::Template;
use Class::Date qw(:errors now);

use lib "/home/fcs/cbryrocher/web/params/";    # permet d'ajouter des repertoires à @INC
use lib "/home/fcs/cbryrocher/web/lib/";       # permet d'ajouter des repertoires à @INC
use conf_cbryrocher;
use lib_cbryrocher;

my $hdl_cgi    = new CGI;
my $ticket     = $hdl_cgi->param('TICKET');
my $date_arrived = $hdl_cgi->param('DATE_ARRIVED');
my $to_excell = $hdl_cgi->param('TO_EXCELL');
my $is_batch   = $hdl_cgi->param('IS_BATCH');
my $methode    = $ENV{REQUEST_METHOD};
my $adresse_ip = $ENV{REMOTE_ADDR};
my $navigateur = $ENV{HTTP_USER_AGENT};

my $uo_like = $hdl_cgi->param('UO_LIKE');
my $date_open_min = $hdl_cgi->param('DATE_OPEN_MIN');
my $date_open_max = $hdl_cgi->param('DATE_OPEN_MAX');

$uo_like = '';

my $date_arrived = '';


my $date_arrived_pod = substr($date_arrived,6,4).substr($date_arrived,3,2).substr($date_arrived,0,2);

my ( $dbh, $acces );
( $dbh = DBI->connect( $base_fcs_dsn, $base_fcs_user, $base_fcs_password, { AutoCommit => 1 } ) ) or &FCS_AfficheErreur( "authentification.psp", "POST", $ticket, "Connexion base impossible." );
  my $date_tmp     = now;
  $date_arrived_pod = $date_tmp->year . sprintf( "%02d", $date_tmp->month ) . sprintf( "%02d", $date_tmp->day );

  
 my $dbh_fcs = DBI->connect( $base_fcs_fcs_dsn , $base_fcs_fcs_user, $base_fcs_fcs_password, { AutoCommit => 1 } ) or &FCS_AfficheErreur( "authentification.psp", "POST", $ticket, "Connexion base impossible." );
 
 my $dbh_cot = DBI->connect( $base_cot_dsn, $base_cot_user, $base_cot_password, { AutoCommit => 1 } );
  
# Pour les UOs Archivé
my $tmp_date = now;
my $date_tmp = $tmp_date->year - 1 . sprintf( "%02d", $tmp_date->month ) . sprintf( "%02d", $tmp_date->day );

my $template = HTML::Template->new( filename => '/home/fcs/cbryrocher/bin/dau_to_excell.tmpl' , global_vars => 1, die_on_bad_params => 0  );

my $DBUG = '0';

my $log_msg;
my $value_to_return;
my $is_print_into_log;

my $uo_num;
my $num_po;


my ( $is_original, $is_checking, $is_embarque, $is_facture );
init();
$uo_like = $uo_num if($uo_num ne '');

my $is_esol=is_user_in_group($dbh,$ticket,"E-SOLUTIONS");
$is_esol = 1 if($ticket eq '');
$DBUG='0' if($is_esol eq '0');
my $DBUG = '0';

if ( $DBUG eq '1' ) {
  open( FICHIER_LOG, "> /home/fcs/cbryrocher/web/dbug.LOG");
  my $date_debut = now;
  print FICHIER_LOG "\n\nDEBUT $date_debut";
}
my $r = shift;

=header

my $sq = "
-- LA COTATION ORIGINALE
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm 
FROM dau_uo du
LEFT JOIN dau_po_root dpr
ON du.id = dpr.id_uo
LEFT JOIN dau_po dp
ON dpr.id = dp.id_po_root
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
LEFT JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
LEFT JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
WHERE dp.id_export_level = '1'
AND TRIM(dbi.invoice_num) = ''
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0'
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')
GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival

-- L'AMENDE (CHECKING) 
UNION
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt ,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm 
FROM dau_uo du
LEFT JOIN dau_po_root dpr
ON du.id = dpr.id_uo
LEFT JOIN dau_po dp
ON dpr.id = dp.id_po_root
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
LEFT JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
LEFT JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
WHERE dp.id_export_level = '2'
AND TRIM(dbi.invoice_num) = ''
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0'
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')
GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival

-- L'ENGAGE (EMBARQUE)
UNION
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt ,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm
FROM dau_uo du
LEFT JOIN dau_po_root dpr
ON du.id = dpr.id_uo
LEFT JOIN dau_po dp
ON dpr.id = dp.id_po_root
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
LEFT JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
LEFT JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
WHERE dp.id_export_level = '3'
AND TRIM(dbi.invoice_num) = ''
-- Supprimé le 16 Novembre 2012, Demande Laetitia Durand, on prend le flottant donc dés que dans le DAU
--AND TRIM(db.date_arrival) < $date_arrived_pod 
--AND TRIM(db.date_arrival) > 0
--AND ( TRIM(substr(db.date_arrival,1,6)) > '200912' OR TRIM(db.date_arrival) = 0 )
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0'
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')
GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival

-- LE FACTURE
UNION
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt ,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm 
FROM dau_uo du
LEFT JOIN dau_po_root dpr
ON du.id = dpr.id_uo
LEFT JOIN dau_po dp
ON dpr.id = dp.id_po_root
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
LEFT JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
LEFT JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
WHERE dp.id_export_level = '3'
AND TRIM(dbi.invoice_num) <> ''
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0' 
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')
GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival
ORDER BY uo,po_root,level,po,bl,invoice
";

=cut


my $sq_originale = "
-- LA COTATION ORIGINALE
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE 
WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm 
FROM dau_uo du
INNER JOIN dau_po_root dpr
ON du.id = dpr.id_uo
INNER JOIN dau_po dp
ON dpr.id = dp.id_po_root
INNER JOIN dau_bl db
ON dp.id = db.id_po
INNER JOIN  dau_bl_invoice dbi
ON db.id = dbi.id_bl
INNER JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
INNER JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
WHERE dp.id_export_level = '1'
AND TRIM(dbi.invoice_num) = ''
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0'
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')

GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival

";

my $sq_checking = "

-- L'AMENDE (CHECKING) 
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE 
WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt ,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm 
FROM dau_uo du
INNER JOIN dau_po_root dpr
ON du.id = dpr.id_uo
INNER JOIN dau_po dp
ON dpr.id = dp.id_po_root
INNER JOIN dau_bl db
ON dp.id = db.id_po
INNER JOIN  dau_bl_invoice dbi
ON db.id = dbi.id_bl
INNER JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
INNER JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
WHERE dp.id_export_level = '2'
AND TRIM(dbi.invoice_num) = ''
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0'
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')

GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival

";
my $sq_embarque = "
-- L'ENGAGE (EMBARQUE)
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE 
WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
MAX(dbi.invoice_amount) as cost,
MAX(ds.import_dpt_fees_rate_by_shpt) as is_import_dpt_fees_rate_by_shpt ,
SUM(ds.sku_pcs) as pcs,
SUM(ds.sku_cbm * ds.sku_ctns) as cbm
FROM dau_uo du
INNER JOIN dau_po_root dpr
ON du.id = dpr.id_uo
INNER JOIN dau_po dp
ON dpr.id = dp.id_po_root
INNER JOIN dau_bl db
ON dp.id = db.id_po
INNER JOIN  dau_bl_invoice dbi
ON db.id = dbi.id_bl
INNER JOIN dau_sku ds
ON ds.id_bl = dbi.id_bl
INNER JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
WHERE dp.id_export_level = '3'
AND TRIM(dbi.invoice_num) = ''
-- Supprimé le 16 Novembre 2012, Demande Laetitia Durand, on prend le flottant donc dés que dans le DAU
--AND TRIM(db.date_arrival) < $date_arrived_pod 
--AND TRIM(db.date_arrival) > 0
--AND ( TRIM(substr(db.date_arrival,1,6)) > '200912' OR TRIM(db.date_arrival) = 0 )
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0'
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')

GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival

";
my $sq_facture = "
-- LE FACTURE
SELECT
TRIM(du.uo_num) as uo,
du.id_status as id_status,
TRIM(rds.designation) as statut,
TRIM(dpr.contact_ada) as contact,
TRIM(dpr.po_root_num) as po_root,
TRIM(dpr.dpt) as dpt,
db.is_not_waiting_invoice as factavenir,
du.date_open as moisouvert,
du.date_resulted as moissolde,
du.date_confirmed as moisconfirm,
du.date_close as moiscloture,
TRIM(dp.po_num) as po,
dp.id_export_level as level,
TRIM(db.bl_num) as bl,
CASE WHEN TRIM(dbi.cost_kind) = 'EGAL DUTIES & TAXES'
        OR TRIM(dbi.cost_kind) = 'DUTIES COSTS'
    OR TRIM(dbi.cost_kind) = 'CUSTOMS FEES' THEN 'CUSTOMS'
WHEN TRIM(dbi.cost_kind) = 'POST ACHEMINEMENT' THEN 'FREIGHT'
ELSE TRIM(dbi.cost_kind)
END as cost_kind,
TRIM(dbi.invoice_num) as invoice,
rdc.designation as currency,
dbi.exchange_rate as rate,
TRIM(db.date_arrival) as date_arrived_pod,
SUM(dbi.invoice_amount) as cost,
ds.max_is_import_dpt_fees_rate_by_shpt as is_import_dpt_fees_rate_by_shpt,
ds.sum_sku_pcs as pcs, 
ds.sum_sku_cbm  as cbm
FROM dau_uo du
INNER JOIN dau_po_root dpr
ON du.id = dpr.id_uo
INNER JOIN dau_po dp
ON dpr.id = dp.id_po_root
INNER JOIN dau_bl db
ON dp.id = db.id_po
INNER JOIN  dau_bl_invoice dbi
ON db.id = dbi.id_bl
INNER JOIN ( 
	SELECT 
    ds.id_bl,
	MAX(ds.import_dpt_fees_rate_by_shpt) as max_is_import_dpt_fees_rate_by_shpt ,
	SUM(ds.sku_pcs) as sum_sku_pcs,
	SUM(ds.sku_cbm * ds.sku_ctns) as sum_sku_cbm
	FROM dau_sku as ds
	GROUP BY ds.id_bl
) as ds
ON db.id = ds.id_bl
INNER JOIN ref_dau_status rds
ON rds.id = du.id_status
LEFT JOIN ref_dau_currency rdc
ON rdc.id = dbi.id_currency
WHERE dp.id_export_level = '3'
AND TRIM(dbi.invoice_num) <> ''
AND du.uo_num IS NOT NULL 
AND du.uo_num <> '0' 
AND du.uo_num <> '23612'
AND du.uo_num LIKE '$uo_like%'
AND TO_CHAR(TO_DATE(du.date_open, 'YYYYMMDD'), 'YYYY') >= TO_CHAR(NOW()- interval '2 year', 'YYYY')

--GROUP BY du.uo_num, du.id_status, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, dp.po_num, dp.id_export_level, db.bl_num, db.date_arrival,dbi.cost_kind, dbi.invoice_num, dbi.exchange_rate, ds.max_is_import_dpt_fees_rate_by_shpt ,ds.sum_sku_pcs , ds.sum_sku_cbm ,rdc.designation, rds.designation
GROUP BY du.uo_num, du.id_status, rds.designation, dpr.contact_ada, dpr.po_root_num, dpr.dpt, db.is_not_waiting_invoice, du.date_open, du.date_resulted, du.date_confirmed, du.date_close, dp.po_num, dp.id_export_level, db.bl_num, dbi.cost_kind, dbi.invoice_num, rdc.designation, dbi.exchange_rate, db.date_arrival, 
ds.max_is_import_dpt_fees_rate_by_shpt, ds.sum_sku_pcs, ds.sum_sku_cbm

ORDER BY uo,po_root,level,po,bl,invoice
";


my $sq;
$sq .= $sq_originale;
$sq .= ' UNION ';
$sq .= $sq_checking;
$sq .= ' UNION ';
$sq .= $sq_embarque;
$sq .= ' UNION ';
$sq .= $sq_facture;

if ( $is_original ne '' ||  $is_checking ne '' || $is_embarque ne '' ||  $is_facture ne '' ) {
	if ( $is_original  ) {
		$sq = $sq_originale;
	}
	if ( $is_checking ne '' ) {
		$sq = $sq_checking;
	}
	if ( $is_embarque ne ''  ) {
		$sq = $sq_embarque;
	}	
	if ( $is_facture ne '' ) {
		$sq = $sq_facture;
	}
}


print $sq if($DBUG);

my $rq=$dbh->prepare($sq);
$rq->execute;
if ( $dbh->errstr ne undef ) {
   print $dbh->errstr.":<br><pre>".$sq; 
   $rq->finish;
   exit;
  }

# les requetes
my $sdec = " SELECT comment FROM dau_extra_cost 
            WHERE id_po_root IN
	    	( SELECT id FROM dau_po_root WHERE po_root_num = ? )
";
my $rdec = $dbh->prepare( $sdec );

my $sdu = " SELECT comment FROM dau_uo 
            WHERE uo_num = ?
			
";
my $rdu = $dbh->prepare( $sdu );

my $sdpr = " SELECT comment FROM dau_po_root 
            WHERE po_root_num = ?
";
my $rdpr = $dbh->prepare( $sdpr );

my $sdp = " SELECT comment FROM dau_po 
            WHERE po_num = ? 
";
my $rdp = $dbh->prepare( $sdp );
    
my $sc1 = " SELECT SUM(dec.extra_cost_amount) 
			FROM dau_po_root dpr
			LEFT JOIN dau_extra_cost dec
            ON dpr.id = dec.id_po_root
            WHERE TRIM(dpr.po_root_num) = ? 
			AND (dec.is_confirmed IS NULL OR dec.is_confirmed <> '1')
            AND (dec.is_closed IS NULL OR dec.is_closed <> '1')
            AND dec.id NOT IN ( SELECT di.id_extra_cost FROM dau_extra_cost_bl_invoice di
            LEFT JOIN dau_extra_cost de
		    ON de.id = di.id_extra_cost )
";
my $rc1 = $dbh->prepare( $sc1 );

my $sc2 = " SELECT SUM(dec.extra_cost_amount) 
			FROM dau_po_root dpr
			LEFT JOIN dau_extra_cost dec
			ON dpr.id = dec.id_po_root
            WHERE TRIM(dpr.po_root_num) = ? 
			AND dec.id IS NOT NULL
			AND (dec.is_confirmed IS NULL OR dec.is_confirmed <> '1')
            AND (dec.is_closed IS NULL OR dec.is_closed <> '1')
            AND dec.id IN ( SELECT di.id_extra_cost FROM dau_extra_cost_bl_invoice di
                       LEFT JOIN dau_extra_cost de
		       ON de.id = di.id_extra_cost )
";
my $rc2 = $dbh->prepare( $sc2 );

my $sc3 = " SELECT SUM(dec.extra_cost_amount) 
            FROM dau_po_root dpr
			LEFT JOIN dau_extra_cost dec
			ON dpr.id = dec.id_po_root
            WHERE TRIM(dpr.po_root_num) = ? 
            AND dec.is_confirmed = '1'
            AND (dec.is_closed IS NULL OR dec.is_closed <> '1')
";
my $rc3 = $dbh->prepare( $sc3 );

my $sc4 = " SELECT SUM(dec.extra_cost_amount) 
            FROM dau_po_root dpr
			LEFT JOIN dau_extra_cost dec
	    ON dpr.id = dec.id_po_root
            WHERE TRIM(dpr.po_root_num) = ? 
	    AND (dec.is_confirmed IS NULL OR dec.is_confirmed <> '1')
            AND dec.is_closed = '1'
			AND dec.id IS NOT NULL
";
my $rc4 = $dbh->prepare( $sc4 );

# POUR COTATION ET AMENDE ON PREND LE PRIX HUB
my $sCAimport = " SELECT 
		TRIM(ds.sku_num) as sku,
		TRIM(ds.sku_size) as size,
		TRIM(ds.sku_color) as color,
		ds.price_hub as prix_hub,
		ds.price_fob1 as prix_fob1,
		ds.price_fob2 as prix_fob2,
		SUM(ds.sku_pcs) as pcs
		FROM dau_po_root dpr
		LEFT JOIN dau_po dp
		ON dpr.id = dp.id_po_root
		LEFT JOIN dau_bl bl
		ON dp.id = bl.id_po
		LEFT JOIN dau_sku ds
		ON bl.id = ds.id_bl
		WHERE dpr.po_root_num = ?
		AND dp.id_export_level = '1'
		GROUP BY ds.sku_num, ds.sku_size, ds.sku_color, ds.price_hub, ds.price_fob1, ds.price_fob2
		ORDER BY ds.sku_num, ds.sku_size, ds.sku_color
";
my $rCAimport = $dbh->prepare( $sCAimport );

# POUR ENGAGE ON PREND LE PRIX FACTURE
my $sCAimportengage = " 
			SELECT
			TRIM(ds.sku_num) as sku,
			TRIM(ds.sku_size) as size,
			TRIM(ds.sku_color) as color,
			df.price_invoiced as prix_facture,
			df.price_fob1 as prix_fob1,
			df.price_fob2 as prix_fob2,
			SUM(df.sku_pcs_deducted) as pcs
			FROM dau_po_root dpr
			LEFT JOIN dau_po dp
			ON dpr.id = dp.id_po_root
			LEFT JOIN dau_bl bl
			ON dp.id = bl.id_po
			LEFT JOIN dau_sku ds
			ON bl.id = ds.id_bl
			LEFT JOIN dau_floating df
			ON ds.id = df.id_sku
			WHERE dpr.po_root_num = ?
			AND df.is_deducted > 0
			AND dp.id_export_level = '3'
-- Supprimé le 16 Novembre 2012, Demande Laetitia Durand, on prend le flottant donc dés que dans le DAU
			--AND TRIM(bl.date_arrival) < $date_arrived_pod 
			--AND TRIM(bl.date_arrival) <> '0'
			AND df.price_invoiced > '0'
			GROUP BY ds.sku_num, ds.sku_size, ds.sku_color, df.price_invoiced, df.price_fob1, df.price_fob2
			ORDER BY ds.sku_num, ds.sku_size, ds.sku_color
";
my $rCAimportengage = $dbh->prepare( $sCAimportengage );

# POUR FACTURE ON PREND LE PRIX FACTURE
my $sCAimportfacture = " SELECT
			TRIM(ds.sku_num) as sku,
			TRIM(ds.sku_size) as size,
			TRIM(ds.sku_color) as color,
			df.price_invoiced as prix_facture,
			df.price_fob1 as prix_fob1,
			df.price_fob2 as prix_fob2,
			SUM(df.sku_pcs_deducted) as pcs
			FROM dau_po_root dpr
			LEFT JOIN dau_po dp
			ON dpr.id = dp.id_po_root
			LEFT JOIN dau_bl bl
			ON dp.id = bl.id_po
			LEFT JOIN dau_sku ds
			ON bl.id = ds.id_bl
			LEFT JOIN dau_floating df
			ON ds.id = df.id_sku
			WHERE dpr.po_root_num = ?
			AND df.is_deducted > 0
			AND dp.id_export_level = '3'
			AND df.price_invoiced > '0'
			GROUP BY ds.sku_num, ds.sku_size, ds.sku_color, df.price_invoiced, df.price_fob1, df.price_fob2
			ORDER BY ds.sku_num, ds.sku_size, ds.sku_color
";
my $rCAimportfacture = $dbh->prepare( $sCAimportfacture );

my $sfeesfacture = " SELECT
			TRIM(ds.sku_num) as sku,
			TRIM(ds.sku_size) as size,
			TRIM(ds.sku_color) as color,
			df.price_invoiced as prix_facture,
			df.price_fob1 as prix_fob1,
			df.price_fob2 as prix_fob2,
			SUM(df.sku_pcs_deducted) as pcs
			FROM dau_po dp
			LEFT JOIN dau_bl bl
			ON dp.id = bl.id_po
			LEFT JOIN dau_sku ds
			ON bl.id = ds.id_bl
			LEFT JOIN dau_floating df
			ON ds.id = df.id_sku
			WHERE dp.po_num = ?
			AND df.is_deducted > 0
			AND dp.id_export_level = '3'
			GROUP BY ds.sku_num, ds.sku_size, ds.sku_color, df.price_invoiced, df.price_fob1, df.price_fob2
			ORDER BY ds.sku_num, ds.sku_size, ds.sku_color
";
my $rfeesfacture = $dbh->prepare( $sfeesfacture );

my $sfeesengage = " SELECT
			TRIM(ds.sku_num) as sku,
			TRIM(ds.sku_size) as size,
			TRIM(ds.sku_color) as color,
			df.price_invoiced as prix_facture,
			df.price_fob1 as prix_fob1,
			df.price_fob2 as prix_fob2,
			SUM(df.sku_pcs_deducted) as pcs
			FROM dau_po dp
			LEFT JOIN dau_bl bl
			ON dp.id = bl.id_po
			LEFT JOIN dau_sku ds
			ON bl.id = ds.id_bl
			LEFT JOIN dau_floating df
			ON ds.id = df.id_sku
			WHERE dp.po_num = ?
			AND df.is_deducted > 0
			AND dp.id_export_level = '3'
-- Supprimé le 16 Novembre 2012, Demande Laetitia Durand, on prend le flottant donc dés que dans le DAU
			--AND bl.date_arrival < $date_arrived_pod
                        --AND TRIM(bl.date_arrival) <> '0'	
			GROUP BY ds.sku_num, ds.sku_size, ds.sku_color, df.price_invoiced, df.price_fob1, df.price_fob2
			ORDER BY ds.sku_num, ds.sku_size, ds.sku_color
";
my $rfeesengage = $dbh->prepare( $sfeesengage );

my $sphf = " 
		SELECT SUM(ds.sku_pcs) as pcs
		FROM dau_po_root dpr
		LEFT JOIN dau_po dp
		ON dpr.id = dp.id_po_root
		LEFT JOIN dau_bl bl
		ON dp.id = bl.id_po
		LEFT JOIN dau_sku ds
		ON bl.id = ds.id_bl
	     WHERE dpr.po_root_num = ?
	     AND dp.id_export_level = ?
	     AND TRIM(ds.sku_num) = ?
	     AND TRIM(ds.sku_size) = ?
	     AND TRIM(ds.sku_color) = ?
";
my $rphf = $dbh->prepare( $sphf );
#CD 20150616 on va chercher les fees de la cotation retenu commerce et non de la cotation Z ou de la cyntir si elle existe...
my $sqlr_fees = " 
SELECT
fic_num,
order_by,

CASE
	WHEN import_fees_rate IS NULL OR import_fees_rate = 0 THEN 0
	ELSE import_fees_rate/100
END as import_fees_rate, 
CASE
	WHEN qc_fees_rate IS NULL OR qc_fees_rate = 0 THEN 0
	ELSE qc_fees_rate/100
END as qc_fees_rate, 
CASE
	WHEN import_risk_fees_rate IS NULL OR import_risk_fees_rate = 0 THEN 0
	ELSE import_risk_fees_rate/100
END as import_risk_fees_rate, 
CASE
	WHEN purchase_fees_rate IS NULL OR purchase_fees_rate = 0 THEN 0
	ELSE purchase_fees_rate/100
END as purchase_fees_rate, 
CASE
	WHEN other_fees_rate IS NULL OR other_fees_rate = 0 THEN 0
	ELSE other_fees_rate/100
END as other_fees_rate, 
CASE
	WHEN exchange_rate_fees IS NULL OR exchange_rate_fees = 0 THEN 0
	ELSE exchange_rate_fees/100
END as exchange_rate_fees


FROM (

SELECT 
'1' as order_by,
fic_num, 
import_dpt_fees as import_fees_rate, 
qc_dpt_fees as qc_fees_rate, 
import_risk_fees as import_risk_fees_rate,
purchase_fees as purchase_fees_rate, 
0 as other_fees_rate, 
exchange_rate_fees as exchange_rate_fees
FROM qf_fees
--WHERE fic_num = SPLIT_PART('4134923G01_4-5', '_', 1)

UNION

SELECT 
'2' as order_by,
SPLIT_PART(MAX(dp1.po_num), '_', 1) as fic_num,
MAX(import_fees_rate) as import_fees_rate, 
MAX(qc_fees_rate) as qc_fees_rate, 
MAX(import_risk_fees_rate) as import_risk_fees_rate, 
MAX(purchase_fees_rate) as purchase_fees_rate, 
MAX(other_fees_rate) as other_fees_rate, 
MAX(exchange_rate_fees) as exchange_rate_fees

FROM dau_sku ds1
LEFT JOIN dau_bl bl1
ON bl1.id = ds1.id_bl
LEFT JOIN dau_po dp1
ON dp1.id = bl1.id_po
AND dp1.id_export_level = '1'
--WHERE SPLIT_PART(dp1.po_num, '_', 1) = SPLIT_PART('4134923G01_4-5', '_', 1)
--WHERE SUBSTR(dp1.po_num,1,10) = SUBSTR('4134923G01_4-5',1,10)
GROUP BY SPLIT_PART(dp1.po_num, '_', 1)

) as main


WHERE 
1=1
AND fic_num = SPLIT_PART(?, '_', 1)
ORDER BY order_by
LIMIT 1 OFFSET 0
;

";

my $rqlr_fees = $dbh->prepare( $sqlr_fees );

my $scontainer_fcs = "

SELECT
main.uo_number, 
main.id_export_level,
type_container_20.nb as nb_20, 
type_container_40.nb as nb_40, 
type_container_40hc.nb as nb_40hc, 
type_container_20fr.nb as nb_20fr, 
type_container_40fr.nb as nb_40fr, 
type_container_40hr.nb as nb_40hr, 
type_container_air.nb as nb_air, 
type_container_lcl.nb as nb_lcl, 
type_container_cfs.nb as nb_cfs, 
''
FROM (
	SELECT
	DISTINCT uo_number, id_export_level
	FROM qf_equipment
) as main
LEFT JOIN (
	SELECT

	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = '20'
	GROUP BY uo_number, id_export_level
) as type_container_20
ON type_container_20.uo_number = main.uo_number
AND type_container_20.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT

	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = '40'
	GROUP BY uo_number, id_export_level
) as type_container_40
ON type_container_40.uo_number = main.uo_number
AND type_container_40.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT

	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = '40HC'
	GROUP BY uo_number, id_export_level
) as type_container_40hc
ON type_container_40hc.uo_number = main.uo_number
AND type_container_40hc.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT
	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = '20FR'
	GROUP BY uo_number, id_export_level
) as type_container_20fr
ON type_container_20fr.uo_number = main.uo_number
AND type_container_20fr.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT
	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = '40FR'
	GROUP BY uo_number, id_export_level
) as type_container_40fr
ON type_container_40fr.uo_number = main.uo_number
AND type_container_40fr.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT
	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = '40HR'
	GROUP BY uo_number, id_export_level
) as type_container_40hr
ON type_container_40hr.uo_number = main.uo_number
AND type_container_40hr.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT
	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = 'AIR'
	GROUP BY uo_number, id_export_level
) as type_container_air
ON type_container_air.uo_number = main.uo_number
AND type_container_air.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT
	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container = 'LCL'
	GROUP BY uo_number, id_export_level
) as type_container_lcl
ON type_container_lcl.uo_number = main.uo_number
AND type_container_lcl.id_export_level = main.id_export_level

LEFT JOIN (
	SELECT
	uo_number,
	id_export_level,
	SUM(qe.nb) as nb
	FROM qf_equipment as qe
	WHERE type_container LIKE 'CFS%'
	GROUP BY uo_number, id_export_level
) as type_container_cfs
ON type_container_cfs.uo_number = main.uo_number
AND type_container_cfs.id_export_level = main.id_export_level

WHERE 1=1
AND main.id_export_level  = ?
AND main.uo_number = ?

";

my $rcontainer_fcs = $dbh->prepare( $scontainer_fcs );

my $secart_engage = "
SELECT
( SELECT SUM(invoice_amount )
FROM dau_po_root dpr
LEFT JOIN dau_po dp
ON dpr.id = dp.id_po_root
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl

WHERE TRIM(dbi.invoice_num) = ''
AND dbi.cost_kind = 'FREIGHT_POL'
AND dp.id_export_level = '3'
AND dp.po_num = ?
) AS cout_reel_engage,
( SELECT SUM(invoice_amount * rate_freight / ( SELECT MAX(rate_freight)
FROM dau_po dp
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
						WHERE TRIM(dbi.invoice_num) = ''
						AND dp.id_export_level IN ('1','2') -- On peut ne pas avoir de cotation initiale et uniquement de l'amendé
						AND dp.po_num =  ? ) 
            )
FROM dau_po dp
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
WHERE TRIM(dbi.invoice_num) = ''
AND dbi.cost_kind = 'FREIGHT_POL'
AND dp.id_export_level = '3'
AND dp.po_num = ?
) AS cout_corrige_engage,
( SELECT SUM(invoice_amount )
FROM dau_po dp
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
WHERE TRIM(dbi.invoice_num) <> ''
AND dbi.cost_kind = 'FREIGHT'
AND dbi.id_issuer = '8' --DHL
AND dp.id_export_level = '3'
AND dp.po_num = ?
) AS cout_reel_facture,
( SELECT SUM(invoice_amount * rate_freight / ( SELECT MAX(rate_freight)
FROM dau_po dp
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
WHERE TRIM(dbi.invoice_num) = ''
AND dp.id_export_level IN ('1','2')
AND dp.po_num =  ? )
	    )
FROM dau_po dp
LEFT JOIN dau_bl db
ON dp.id = db.id_po
LEFT JOIN dau_bl_invoice dbi
ON db.id = dbi.id_bl
WHERE TRIM(dbi.invoice_num) <> ''
AND dbi.cost_kind = 'FREIGHT'
AND dbi.id_issuer = '8' --DHL
AND dp.id_export_level = '3'
AND dp.po_num = ?
) AS cout_corrige_facture
;
";

my $recart_engage = $dbh->prepare( $secart_engage );

my @loop;
my $level = '';
my $purchase_fees=0.0;
my $is_import_dpt_fees_rate_by_shpt;
my ( $UO, $STATUT, $CONTACT, $PO_ROOT, $FACTAVENIR, $MOISOUVERT, $MOISSOLDE, $MOISCONFIRM, $MOISCLOTURE, $PO, $BL, $PCS ) = '';
my ( $shpt_cot, $shpt_amende, $shpt_bl, $shpt_facture ) = '';
my ( $c_COST_FREIGHT, $c_COST_PORT, $c_COST_CUSTOM, $c_COST_FINAL, $c_COST_WHSE, $c_COST_FEES, $c_PURCHASE_FEES,$c_COST_ADMIN, $c_COST_OTHER, $c_total_cost ) = '';
my ( $a_COST_FREIGHT, $a_COST_PORT, $a_COST_CUSTOM, $a_COST_FINAL, $a_COST_WHSE, $a_COST_FEES, $a_PURCHASE_FEES,$a_COST_ADMIN, $a_COST_OTHER, $a_total_cost ) = '';
my ( $e_COST_FREIGHT, $e_COST_PORT, $e_COST_CUSTOM, $e_COST_FINAL, $e_COST_WHSE, $e_COST_FEES, $e_PURCHASE_FEES,$e_COST_ADMIN, $e_COST_OTHER, $e_total_cost ) = '';
my ( $f_COST_FREIGHT, $f_COST_PORT, $f_COST_CUSTOM, $f_COST_FINAL, $f_COST_WHSE, $f_COST_FEES, $f_PURCHASE_FEES,$f_COST_ADMIN, $f_COST_OTHER, $f_total_cost ) = '';
my ( $c_CURRENCY, $c_RATE, $a_CURRENCY, $a_RATE, $f_CURRENCY, $f_RATE, $e_CURRENCY, $e_RATE ) = '';
my $key = '';
my ( $old_bl, $old_po, $old_invoice, $pcs_cot, $pcs_amende, $pcs_engage ) = ''; 
my ( $cbm_cot, $cbm_amende, $cbm_engage ) = ''; 
my ( $ecart_engage, $ecart_facture ) = '';

while (  my $data = $rq->fetchrow_hashref ) {
  if ( $key eq '') {
    $key = $data->{'uo'}.'_'.$data->{'statut'}.'_'.$data->{'contact'}.'_'.$data->{'po_root'}.'_'.$data->{'factavenir'}.'_'.$data->{'moisouvert'}.'_'.$data->{'moissolde'}.'_'.$data->{'moisconfirm'}.'_'.$data->{'moiscloture'};
  }
  if ( $key ne $data->{'uo'}.'_'.$data->{'statut'}.'_'.$data->{'contact'}.'_'.$data->{'po_root'}.'_'.$data->{'factavenir'}.'_'.$data->{'moisouvert'}.'_'.$data->{'moissolde'}.'_'.$data->{'moisconfirm'}.'_'.$data->{'moiscloture'} ) {
    my %hash = ();
    $hash{UO} = $UO;
    $hash{STATUT} = $STATUT;
    $hash{CONTACT} = $CONTACT;
    if ( $FACTAVENIR eq '0' ) { $FACTAVENIR = 'Oui'; }
    else { $FACTAVENIR = 'Non'; }
    $hash{FACTAVENIR} = $FACTAVENIR;
    $hash{MOISOUVERT} = substr($MOISOUVERT,0,4).'-'.substr($MOISOUVERT,4,2);
    $hash{MOISSOLDE} = substr($MOISSOLDE,0,4).'-'.substr($MOISSOLDE,4,2);
    $hash{MOISCONFIRM} = substr($MOISCONFIRM,0,4).'-'.substr($MOISCONFIRM,4,2);
    $hash{MOISCLOTURE} = substr($MOISCLOTURE,0,4).'-'.substr($MOISCLOTURE,4,2);
    $hash{PO_ROOT} = $PO_ROOT;
    $hash{SHPT_COT} = $shpt_cot;
    $hash{SHPT_AMENDE} = $shpt_amende;
    $hash{SHPT_BL} = $shpt_bl;
    $hash{SHPT_FACTURE} = $shpt_facture;
	$hash{PURCHASE_RATE} = $purchase_fees;
	$hash{IS_ID_FEES_RATE_BY_SHPT}= 'X' if($is_import_dpt_fees_rate_by_shpt ne '');
    $hash{C_FREIGHT} = sprintf ("%0.1f", $c_COST_FREIGHT);
    $hash{C_PORT} = sprintf ("%0.1f", $c_COST_PORT);
    $hash{C_CUSTOM} = sprintf ("%0.1f", $c_COST_CUSTOM);
    $hash{C_FINAL} = sprintf ("%0.1f", $c_COST_FINAL);
    $hash{C_WHSE} = sprintf ("%0.1f", $c_COST_WHSE);
    $hash{C_FEES} = sprintf ("%0.1f", $c_COST_FEES);
	$hash{C_PURCHASE_FEES} = sprintf ("%0.1f", $c_PURCHASE_FEES);
    $hash{C_ADMIN} = sprintf ("%0.1f", $c_COST_ADMIN);
    $hash{C_OTHER} = sprintf ("%0.1f", $c_COST_OTHER);
    $hash{C_TOTAL} = sprintf ("%0.1f", $c_total_cost);
    $hash{C_CURRENCY} = $c_CURRENCY;
    $hash{C_RATE} = $c_RATE;
	
# Recherche des containers

	$hash{C_LCL} = '';
	$hash{C_20} = '';
	$hash{C_20FR} = '';
	$hash{C_40} = '';
	$hash{C_40FR} = '';
	$hash{C_45} = '';
	$hash{C_40HC} = '';
	$hash{C_40HR} = '';
	$hash{C_CFS} = '';
	$hash{C_AIR} = '';



   $rcontainer_fcs->execute(1,$UO);


	while (my $container = $rcontainer_fcs->fetchrow_hashref) {

	  $hash{C_LCL} = $container->{'nb_lcl'}; 
	  $hash{C_20} = $container->{'nb_20'}; 
	  $hash{C_40} = $container->{'nb_40'}; 
	  $hash{C_45} =  $container->{'nb_45'};  
	  $hash{C_40HC} = $container->{'nb_40hc'}; 
	  $hash{C_20FR} = $container->{'nb_20fr'}; 
	  $hash{C_40FR} = $container->{'nb_40fr'};
	  $hash{C_40HR} = $container->{'nb_40hr'};
	  $hash{C_CFS} = $container->{'nb_cfs'};
	  $hash{C_AIR} = $container->{'nb_air'};
    }

	$rcontainer_fcs->finish;


    $hash{A_FREIGHT} = sprintf ("%0.1f", $a_COST_FREIGHT);
    $hash{A_PORT} = sprintf ("%0.1f", $a_COST_PORT);
    $hash{A_CUSTOM} = sprintf ("%0.1f", $a_COST_CUSTOM);
    $hash{A_FINAL} = sprintf ("%0.1f", $a_COST_FINAL);
    $hash{A_WHSE} = sprintf ("%0.1f", $a_COST_WHSE);
    $hash{A_FEES} = sprintf ("%0.1f", $a_COST_FEES);
	$hash{A_PURCHASE_FEES} = sprintf ("%0.1f", $a_PURCHASE_FEES);
    $hash{A_ADMIN} = sprintf ("%0.1f", $a_COST_ADMIN);
    $hash{A_OTHER} = sprintf ("%0.1f", $a_COST_OTHER);
    $hash{A_TOTAL} = sprintf ("%0.1f", $a_total_cost);
    $hash{A_CURRENCY} = $a_CURRENCY;
    $hash{A_RATE} = $a_RATE;
	
# Recherche des containers

	$hash{A_LCL} = '';
	$hash{A_20} = '';
	$hash{A_20FR} = '';
	$hash{A_40} = '';
	$hash{A_40FR} = '';
	$hash{A_45} = '';
	$hash{A_40HC} = '';
	$hash{A_40HR} = '';
	$hash{A_CFS} = '';
	$hash{A_AIR} = '';

#$rcontainer->execute($UO, $PO_ROOT, '2', $UO, $PO_ROOT, '2');
$rcontainer_fcs->execute('2',$UO);

	while (my $container = $rcontainer_fcs->fetchrow_hashref) {

	  $hash{A_LCL} = $container->{'nb_lcl'}; 
	  $hash{A_20} = $container->{'nb_20'}; 
	  $hash{A_40} = $container->{'nb_40'}; 
	  $hash{A_45} =  $container->{'nb_45'};  
	  $hash{A_40HC} = $container->{'nb_40hc'}; 
	  $hash{A_20FR} = $container->{'nb_20fr'}; 
	  $hash{A_40FR} = $container->{'nb_40fr'};
	  $hash{A_40HR} = $container->{'nb_40hr'};
	  $hash{A_CFS} = $container->{'nb_cfs'};
	  $hash{A_AIR} = $container->{'nb_air'};
    }

	$rcontainer_fcs->finish;



    $hash{F_FREIGHT} = sprintf ("%0.1f", $f_COST_FREIGHT);
    $hash{F_PORT} = sprintf ("%0.1f", $f_COST_PORT);
    $hash{F_CUSTOM} = sprintf ("%0.1f", $f_COST_CUSTOM);
    $hash{F_FINAL} = sprintf ("%0.1f", $f_COST_FINAL);
    $hash{F_WHSE} = sprintf ("%0.1f", $f_COST_WHSE);
    $hash{F_FEES} = sprintf ("%0.1f", $f_COST_FEES);
	$hash{F_PURCHASE_FEES} = sprintf ("%0.1f", $f_PURCHASE_FEES);
    $hash{F_ADMIN} = sprintf ("%0.1f", $f_COST_ADMIN);
    $hash{F_OTHER} = sprintf ("%0.1f", $f_COST_OTHER);
    $f_total_cost = $f_total_cost + $f_COST_FEES;
    $hash{F_TOTAL} = sprintf ("%0.1f", $f_total_cost);
    $hash{F_CURRENCY} = $f_CURRENCY;
    $hash{F_RATE} = $f_RATE;
    if ( $FACTAVENIR eq 'Oui' ) {
      $hash{E_FREIGHT} = sprintf ("%0.1f", $e_COST_FREIGHT);
      $hash{E_PORT} = sprintf ("%0.1f", $e_COST_PORT);
      $hash{E_CUSTOM} = sprintf ("%0.1f", $e_COST_CUSTOM);
      $hash{E_FINAL} = sprintf ("%0.1f", $e_COST_FINAL);
      $hash{E_WHSE} = sprintf ("%0.1f", $e_COST_WHSE);
      $hash{E_FEES} = sprintf ("%0.1f", $e_COST_FEES);
	  $hash{E_PURCHASE_FEES} = sprintf ("%0.1f", $e_PURCHASE_FEES);
      $hash{E_ADMIN} = sprintf ("%0.1f", $e_COST_ADMIN);
      $hash{E_OTHER} = sprintf ("%0.1f", $e_COST_OTHER);
      $e_total_cost = $e_total_cost + $e_COST_FEES;
      $hash{E_TOTAL} = sprintf ("%0.1f", $e_total_cost);
      $hash{E_CURRENCY} = $e_CURRENCY;
      $hash{E_RATE} = $e_RATE;
    }
    else {
      $hash{E_FREIGHT} = sprintf ("%0.1f", $f_COST_FREIGHT);
      $hash{E_PORT} = sprintf ("%0.1f", $f_COST_PORT);
      $hash{E_CUSTOM} = sprintf ("%0.1f", $f_COST_CUSTOM);
      $hash{E_FINAL} = sprintf ("%0.1f", $f_COST_FINAL);
      $hash{E_WHSE} = sprintf ("%0.1f", $f_COST_WHSE);
      $hash{E_FEES} = sprintf ("%0.1f", $f_COST_FEES);
	  $hash{E_PURCHASE_FEES} = sprintf ("%0.1f", $f_PURCHASE_FEES);
      $hash{E_ADMIN} = sprintf ("%0.1f", $f_COST_ADMIN);
      $hash{E_OTHER} = sprintf ("%0.1f", $f_COST_OTHER);
      $hash{E_TOTAL} = sprintf ("%0.1f", $f_total_cost);
      $hash{E_CURRENCY} = $f_CURRENCY;
      $hash{E_RATE} = $f_RATE;
    }
	
	# Recherche des containers
	$hash{E_LCL} = '';
	$hash{E_20} = '';
	$hash{E_20FR} = '';
	$hash{E_40} = '';
	$hash{E_40FR} = '';
	$hash{E_45} = '';
	$hash{E_40HC} = '';
	$hash{E_40HR} = '';
	$hash{E_CFS} = '';
	$hash{E_AIR} = '';

	$rcontainer_fcs->execute('3',$UO);

	while (my $container = $rcontainer_fcs->fetchrow_hashref) {

	  $hash{E_LCL} = $container->{'nb_lcl'}; 
	  $hash{E_20} = $container->{'nb_20'}; 
	  $hash{E_40} = $container->{'nb_40'}; 
	  $hash{E_45} =  $container->{'nb_45'};  
	  $hash{E_40HC} = $container->{'nb_40hc'}; 
	  $hash{E_20FR} = $container->{'nb_20fr'}; 
	  $hash{E_40FR} = $container->{'nb_40fr'};
	  $hash{E_40HR} = $container->{'nb_40hr'};
	  $hash{E_CFS} = $container->{'nb_cfs'};
	  $hash{E_AIR} = $container->{'nb_air'};
    }

	$rcontainer_fcs->finish;

    $hash{ECART_AC_FREIGHT} = sprintf ("%0.1f", $a_COST_FREIGHT - $c_COST_FREIGHT);
    $hash{ECART_AC_PORT} = sprintf ("%0.1f", $a_COST_PORT - $c_COST_PORT);
    $hash{ECART_AC_CUSTOM} = sprintf ("%0.1f", $a_COST_CUSTOM - $c_COST_CUSTOM);
    $hash{ECART_AC_FINAL} = sprintf ("%0.1f", $a_COST_FINAL - $c_COST_FINAL);
    $hash{ECART_AC_WHSE} = sprintf ("%0.1f", $a_COST_WHSE - $c_COST_WHSE);
    $hash{ECART_AC_FEES} = sprintf ("%0.1f", $a_COST_FEES - $c_COST_FEES);
    $hash{ECART_AC_ADMIN} = sprintf ("%0.1f", $a_COST_ADMIN - $c_COST_ADMIN);
    $hash{ECART_AC_OTHER} = sprintf ("%0.1f", $a_COST_OTHER - $c_COST_OTHER);
    $hash{TOTAL_ECART_AC} = sprintf ("%0.1f", $a_total_cost - $c_total_cost);
    $hash{ECART_AF_FREIGHT} = sprintf ("%0.1f", $f_COST_FREIGHT - $c_COST_FREIGHT);
    $hash{ECART_AF_PORT} = sprintf ("%0.1f", $f_COST_PORT - $c_COST_PORT);
    $hash{ECART_AF_CUSTOM} = sprintf ("%0.1f", $f_COST_CUSTOM - $c_COST_CUSTOM);
    $hash{ECART_AF_FINAL} = sprintf ("%0.1f", $f_COST_FINAL - $c_COST_FINAL);
    $hash{ECART_AF_WHSE} = sprintf ("%0.1f", $f_COST_WHSE - $c_COST_WHSE);
    $hash{ECART_AF_FEES} = sprintf ("%0.1f", $f_COST_FEES - $c_COST_FEES);
    $hash{ECART_AF_ADMIN} = sprintf ("%0.1f", $f_COST_ADMIN - $c_COST_ADMIN);
    $hash{ECART_AF_OTHER} = sprintf ("%0.1f", $f_COST_OTHER - $c_COST_OTHER);
    $hash{TOTAL_ECART_AF} = sprintf ("%0.1f", $f_total_cost - $c_total_cost);
    if ( $FACTAVENIR eq 'Oui' ) {
      $hash{ECART_EF_FREIGHT} = sprintf ("%0.1f", $e_COST_FREIGHT - $f_COST_FREIGHT);
      $hash{ECART_EF_PORT} = sprintf ("%0.1f", $e_COST_PORT - $f_COST_PORT);
      $hash{ECART_EF_CUSTOM} = sprintf ("%0.1f", $e_COST_CUSTOM - $f_COST_CUSTOM);
      $hash{ECART_EF_FINAL} = sprintf ("%0.1f", $e_COST_FINAL - $f_COST_FINAL);
      $hash{ECART_EF_WHSE} = sprintf ("%0.1f", $e_COST_WHSE - $f_COST_WHSE);
      $hash{ECART_EF_FEES} = sprintf ("%0.1f", $e_COST_FEES - $f_COST_FEES);
      $hash{ECART_EF_ADMIN} = sprintf ("%0.1f", $e_COST_ADMIN - $f_COST_ADMIN);
      $hash{ECART_EF_OTHER} = sprintf ("%0.1f", $e_COST_OTHER - $f_COST_OTHER);
      $hash{TOTAL_ECART_EF} = sprintf ("%0.1f", $e_total_cost - $f_total_cost);
    }
    else {
      $hash{ECART_EF_FREIGHT} = '0' ;
      $hash{ECART_EF_PORT} = '0';
      $hash{ECART_EF_CUSTOM} = '0';
      $hash{ECART_EF_FINAL} = '0';
      $hash{ECART_EF_WHSE} = '0';
      $hash{ECART_EF_FEES} = '0';
      $hash{ECART_EF_ADMIN} = '0';
      $hash{ECART_EF_OTHER} = '0';
      $hash{TOTAL_ECART_EF} = '0';
    }
    if ( $pcs_cot == 0 ) { $pcs_cot = 1; }
    my $bis_freight = $c_COST_FREIGHT / $pcs_cot * $pcs_engage;
    my $bis_port = $c_COST_PORT / $pcs_cot * $pcs_engage;
    my $bis_custom = $c_COST_CUSTOM / $pcs_cot * $pcs_engage;
    my $bis_final = $c_COST_FINAL / $pcs_cot * $pcs_engage;
    my $bis_whse = $c_COST_WHSE / $pcs_cot * $pcs_engage;
    my $bis_fees = $c_COST_FEES / $pcs_cot * $pcs_engage;
    my $bis_admin = $c_COST_ADMIN / $pcs_cot * $pcs_engage;
    my $bis_other = $c_COST_OTHER / $pcs_cot * $pcs_engage;
    my $bis_total = $c_total_cost / $pcs_cot * $pcs_engage;
    if ( $FACTAVENIR eq 'Oui' ) {
      $hash{ECART_BE_FREIGHT} = sprintf ("%0.1f", $bis_freight - $e_COST_FREIGHT);
      $hash{ECART_BE_PORT} = sprintf ("%0.1f", $bis_port - $e_COST_PORT);
      $hash{ECART_BE_CUSTOM} = sprintf ("%0.1f", $bis_custom - $e_COST_CUSTOM);
      $hash{ECART_BE_FINAL} = sprintf ("%0.1f", $bis_final - $e_COST_FINAL);
      $hash{ECART_BE_WHSE} = sprintf ("%0.1f", $bis_whse - $e_COST_WHSE);
      $hash{ECART_BE_FEES} = sprintf ("%0.1f", $bis_fees - $e_COST_FEES);
      $hash{ECART_BE_ADMIN} = sprintf ("%0.1f", $bis_admin - $e_COST_ADMIN);
      $hash{ECART_BE_OTHER} = sprintf ("%0.1f", $bis_other - $e_COST_OTHER);
      $hash{TOTAL_ECART_BE} = sprintf ("%0.1f", $bis_total - $e_total_cost);
    }
    else {
      $hash{ECART_BE_FREIGHT} = sprintf ("%0.1f", $bis_freight - $f_COST_FREIGHT);
      $hash{ECART_BE_PORT} = sprintf ("%0.1f", $bis_port - $f_COST_PORT);
      $hash{ECART_BE_CUSTOM} = sprintf ("%0.1f", $bis_custom - $f_COST_CUSTOM);
      $hash{ECART_BE_FINAL} = sprintf ("%0.1f", $bis_final - $f_COST_FINAL);
      $hash{ECART_BE_WHSE} = sprintf ("%0.1f", $bis_whse - $f_COST_WHSE);
      $hash{ECART_BE_FEES} = sprintf ("%0.1f", $bis_fees - $f_COST_FEES);
      $hash{ECART_BE_ADMIN} = sprintf ("%0.1f", $bis_admin - $f_COST_ADMIN);
      $hash{ECART_BE_OTHER} = sprintf ("%0.1f", $bis_other - $f_COST_OTHER);
      $hash{TOTAL_ECART_BE} = sprintf ("%0.1f", $bis_total - $f_total_cost);
    }
    $hash{PCS_COT} = $pcs_cot;
    $hash{PCS_AMENDE} = $pcs_amende;
    $hash{PCS_ENGAGE} = $pcs_engage;
    $hash{CBM_COT} = sprintf("%0.2f",$cbm_cot);
    $hash{CBM_AMENDE} = sprintf("%0.2f",$cbm_amende);
    $hash{CBM_ENGAGE} = sprintf("%0.2f",$cbm_engage);
    
    # CA Import Net
    my $c_caimportnet = 0;
    my $a_caimportnet = 0;
    $rCAimport->execute( $PO_ROOT );
    while ( my $tabca = $rCAimport->fetchrow_hashref ) {
      $c_caimportnet = $c_caimportnet + ( $tabca->{'pcs'} * ($tabca->{'prix_hub'} - $tabca->{'prix_fob2'}) ); 
      $rphf->execute( $PO_ROOT, '2', $tabca->{'sku'},  $tabca->{'size'},  $tabca->{'color'}); 
      while ( my $tabsku = $rphf->fetchrow_hashref) {
        $a_caimportnet = $a_caimportnet + ( $tabsku->{'pcs'} * ($tabca->{'prix_hub'} - $tabca->{'prix_fob2'}) );
      }
      $rphf->finish;
    }
    $rCAimport->finish;
    $hash{C_CAIMPORTNET} = sprintf ("%0.1f", $c_caimportnet - $c_COST_FEES);
    $hash{A_CAIMPORTNET} = sprintf ("%0.1f", $a_caimportnet - $a_COST_FEES);

    # POUR l'engagé
    ################
    my $e_caimportnet = 0;
    $rCAimportengage->execute( $PO_ROOT );
    while ( my $tabce = $rCAimportengage->fetchrow_hashref ) {
      $e_caimportnet = $e_caimportnet + ( $tabce->{'pcs'} * ($tabce->{'prix_facture'} - $tabce->{'prix_fob2'}) ); 
    }
    $rCAimportengage->finish;
    $hash{E_CAIMPORTNET} = sprintf ("%0.1f", $e_caimportnet - $e_COST_FEES);
    
    # POUR le facturé
    ################
    my $f_caimportnet = 0;
    $rCAimportfacture->execute( $PO_ROOT );
    while ( my $tabce = $rCAimportfacture->fetchrow_hashref ) {
      $f_caimportnet = $f_caimportnet + ( $tabce->{'pcs'} * ($tabce->{'prix_facture'} - $tabce->{'prix_fob2'}) ); 
    }
    $rCAimportfacture->finish;
    $hash{F_CAIMPORTNET} = sprintf ("%0.1f", $f_caimportnet - $f_COST_FEES);

    # On recherche les surcouts 
    $rc1->execute( $PO_ROOT );
    my $surcout_ouvert = $rc1->fetchrow;
    $rc1->finish;
    $hash{SURCOUTOUVERT} = $surcout_ouvert;
    $rc2->execute( $PO_ROOT );
    my $surcout_rapproche = $rc2->fetchrow;
    $rc2->finish;
    $hash{SURCOUTRAPPROCHE} = $surcout_rapproche;
    $rc3->execute( $PO_ROOT );
    my $surcout_confirme = $rc3->fetchrow;
    $rc3->finish;
    $hash{SURCOUTCONFIRME} = $surcout_confirme;
    $rc4->execute( $PO_ROOT );
    my $surcout_cloture = $rc4->fetchrow;
    $rc4->finish;
    $hash{SURCOUTCLOTURE} = $surcout_cloture;
    $hash{TOTALSURCOUT} = $surcout_ouvert + $surcout_rapproche + $surcout_confirme + $surcout_cloture ;
    
    my $commentsurcout = '';
    $rdec->execute( $PO_ROOT );
    while ( my @result = $rdec->fetchrow ) {
      $commentsurcout = $commentsurcout . $result[0] . ' ';
    }
    $rdec->finish;
    $hash{COMMENTSURCOUT} = $commentsurcout;

    $hash{COMMENTCA} = '';

    my $commentop = '';
    $rdu->execute( $UO );
    while ( my @result = $rdu->fetchrow ) {
      $commentop = $commentop . $result[0] . ' ';
    }
    $rdu->finish;
    $hash{COMMENTOP} = $commentop;

    my $commentporoot = '';
    $rdpr->execute( $PO_ROOT );
    while ( my @result = $rdpr->fetchrow ) {
      $commentporoot = $commentporoot . $result[0] . ' ';
    }
    $rdpr->finish;
    $hash{COMMENTPOROOT} = $commentporoot;

    my $commentpo = '';
    $rdp->execute( $PO );
    while ( my @result = $rdp->fetchrow ) {
      $commentpo = $commentpo . $result[0] . ' ';
    }
    $rdp->finish;
    $hash{COMMENTPO} = $commentpo;

    #$hash{C_RESULTATOP} = $hash{C_CAIMPORTNET} - $hash{C_TOTAL};
    #$hash{A_RESULTATOP} = $hash{A_CAIMPORTNET} - $hash{A_TOTAL};
    #$hash{F_RESULTATOP} = $hash{F_CAIMPORTNET} - $hash{F_TOTAL} - $hash{SURCOUTOUVERT};
    #$hash{E_RESULTATOP} = $hash{E_CAIMPORTNET} - $hash{E_TOTAL};
    # Les Fees sont déjà déduites dans le CA Import Net
    $hash{C_RESULTATOP} = $hash{C_CAIMPORTNET} - $hash{C_TOTAL} + $hash{C_FEES};
    $hash{A_RESULTATOP} = $hash{A_CAIMPORTNET} - $hash{A_TOTAL} + $hash{A_FEES};
    $hash{F_RESULTATOP} = $hash{F_CAIMPORTNET} - $hash{F_TOTAL} + $hash{F_FEES} - $hash{SURCOUTOUVERT};
    $hash{E_RESULTATOP} = $hash{E_CAIMPORTNET} - $hash{E_TOTAL} + $hash{E_FEES};
    $hash{ECART_ENGAGE} = sprintf ("%0.1f",$ecart_engage);
    $hash{ECART_FACTURE} = sprintf ("%0.1f",$ecart_facture);

    push ( @loop, \%hash );

    ( $UO, $STATUT, $CONTACT, $PO_ROOT, $FACTAVENIR, $MOISOUVERT, $MOISSOLDE, $MOISCONFIRM, $MOISCLOTURE, $PO, $BL ,$PCS) = '';
    ( $shpt_cot, $shpt_amende, $shpt_bl, $shpt_facture ) = '';
    ( $c_COST_FREIGHT, $c_COST_PORT, $c_COST_CUSTOM, $c_COST_FINAL, $c_COST_WHSE, $c_COST_FEES, $c_COST_ADMIN, $c_COST_OTHER, $c_total_cost ) = '';
    ( $a_COST_FREIGHT, $a_COST_PORT, $a_COST_CUSTOM, $a_COST_FINAL, $a_COST_WHSE, $a_COST_FEES, $a_COST_ADMIN, $a_COST_OTHER, $a_total_cost ) = '';
    ( $e_COST_FREIGHT, $e_COST_PORT, $e_COST_CUSTOM, $e_COST_FINAL, $e_COST_WHSE, $e_COST_FEES, $e_COST_ADMIN, $e_COST_OTHER, $e_total_cost ) = '';
    ( $f_COST_FREIGHT, $f_COST_PORT, $f_COST_CUSTOM, $f_COST_FINAL, $f_COST_WHSE, $f_COST_FEES, $f_COST_ADMIN, $f_COST_OTHER, $f_total_cost ) = '';
    ( $c_CURRENCY, $c_RATE, $a_CURRENCY, $a_RATE, $f_CURRENCY, $f_RATE, $e_CURRENCY, $e_RATE ) = '';
    ( $old_bl, $old_po, $old_invoice, $pcs_cot, $pcs_amende, $pcs_engage ) = ''; 
    ( $cbm_cot, $cbm_amende, $cbm_engage ) = ''; 
    ( $ecart_engage, $ecart_facture ) = '';
    $key = $data->{'uo'}.'_'.$data->{'statut'}.'_'.$data->{'contact'}.'_'.$data->{'po_root'}.'_'.$data->{'factavenir'}.'_'.$data->{'moisouvert'}.'_'.$data->{'moissolde'}.'_'.$data->{'moisconfirm'}.'_'.$data->{'moiscloture'};
  }

  $UO = $data->{'uo'};
  $STATUT = $data->{'statut'};
  if ( $data->{'id_status'} == '50' && $data->{'moiscloture'} <= $date_tmp ) {
    $STATUT = 'Archivé';
  } 
  $CONTACT = $data->{'contact'};
  $FACTAVENIR = $data->{'factavenir'};
  $MOISOUVERT = $data->{'moisouvert'};
  $MOISSOLDE = $data->{'moissolde'};
  $MOISCONFIRM = $data->{'moisconfirm'};
  $MOISCLOTURE = $data->{'moiscloture'};
  $PO_ROOT = $data->{'po_root'};
  $PO = $data->{'po'};
  $is_import_dpt_fees_rate_by_shpt=$data->{'is_import_dpt_fees_rate_by_shpt'};

  # LA COTATION ORIGINALE
  if ( $data->{'level'} eq '1' ) { 
    if ( $old_bl ne $data->{'bl'} ) {
      $shpt_cot = $shpt_cot . $data->{'bl'} . ';'; 
      $pcs_cot = $pcs_cot + $data->{'pcs'};
      $cbm_cot = $cbm_cot + $data->{'cbm'};
      $old_bl = $data->{'bl'};
    }
    $c_CURRENCY = $data->{'currency'};
    $c_RATE = $data->{'rate'};
    if ( $data->{'cost_kind'} eq 'FREIGHT' ) { $c_COST_FREIGHT = $c_COST_FREIGHT + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'PORT SERVICES' ) { $c_COST_PORT = $c_COST_PORT + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'CUSTOMS' ) { $c_COST_CUSTOM = $c_COST_CUSTOM + $data->{'cost'}; }
    if ( substr($data->{'dpt'},0,4) eq 'Ind_' ) {
      if ( $data->{'cost_kind'} eq 'FINAL DELIVERY' ) { $c_COST_FINAL = $c_COST_FINAL + $data->{'cost'}; }
      if ( $data->{'cost_kind'} eq 'WAREHOUSING' ) { $c_COST_WHSE = $c_COST_WHSE + $data->{'cost'}; }
    }
    if ( 	
		$data->{'cost_kind'} eq 'QC DPT FEES' 
		|| $data->{'cost_kind'} eq 'IMPORT DPT FEES' 
		|| $data->{'cost_kind'} eq 'IMPORT RISK FEES' 
		|| $data->{'cost_kind'} eq 'PURCHASE FEES' 
		|| $data->{'cost_kind'} eq 'OTHER FEES' 
	) { $c_COST_FEES = $c_COST_FEES + $data->{'cost'}; 	}
    if ( $data->{'cost_kind'} eq 'ADMINISTRATIVE LOADS' ) { $c_COST_ADMIN = $c_COST_ADMIN + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'OTHER COST' ) { $c_COST_OTHER = $c_COST_OTHER + $data->{'cost'}; }
    if ( substr($data->{'dpt'},0,4) ne 'Ind_' && ($data->{'cost_kind'} eq 'FINAL DELIVERY' || $data->{'cost_kind'} eq 'WAREHOUSING') ) {}
    else {
      $c_total_cost = $c_total_cost + $data->{'cost'};
    }
  }
  
  # L'AMENDE
  if ( $data->{'level'} eq '2' ) { 
    if ( $old_bl ne $data->{'bl'} ) {
      $shpt_amende = $shpt_amende . $data->{'bl'} . ';'; 
      $pcs_amende = $pcs_amende + $data->{'pcs'}; 
      $cbm_amende = $cbm_amende + $data->{'cbm'}; 
      $old_bl = $data->{'bl'};
    }
    $a_CURRENCY = $data->{'currency'};
    $a_RATE = $data->{'rate'};
    if ( $data->{'cost_kind'} eq 'FREIGHT' ) { $a_COST_FREIGHT = $a_COST_FREIGHT + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'PORT SERVICES' ) { $a_COST_PORT = $a_COST_PORT + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'CUSTOMS' ) { $a_COST_CUSTOM = $a_COST_CUSTOM + $data->{'cost'}; }
    if ( substr($data->{'dpt'},0,4) eq 'Ind_' ) {
      if ( $data->{'cost_kind'} eq 'FINAL DELIVERY' ) { $a_COST_FINAL = $a_COST_FINAL + $data->{'cost'}; }
      if ( $data->{'cost_kind'} eq 'WAREHOUSING' ) { $a_COST_WHSE = $a_COST_WHSE + $data->{'cost'}; }
    }
    if ( 
		$data->{'cost_kind'} eq 'QC DPT FEES' 
		|| $data->{'cost_kind'} eq 'IMPORT DPT FEES' 
		|| $data->{'cost_kind'} eq 'IMPORT RISK FEES' 
		|| $data->{'cost_kind'} eq 'PURCHASE FEES' 
		|| $data->{'cost_kind'} eq 'OTHER FEES' 
	) { $a_COST_FEES = $a_COST_FEES + $data->{'cost'}; }
	if ( $data->{'cost_kind'} eq 'PURCHASE FEES' ) { $a_PURCHASE_FEES = $a_PURCHASE_FEES + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'ADMINISTRATIVE LOADS' ) { $a_COST_ADMIN = $a_COST_ADMIN + $data->{'cost'}; }
    if ( $data->{'cost_kind'} eq 'OTHER COST' ) { $a_COST_OTHER = $a_COST_OTHER + $data->{'cost'}; }
    if ( substr($data->{'dpt'},0,4) ne 'Ind_' && ($data->{'cost_kind'} eq 'FINAL DELIVERY' || $data->{'cost_kind'} eq 'WAREHOUSING') ) {}
    else {
      $a_total_cost = $a_total_cost + $data->{'cost'};
    }
  }

  # LE FACTURE
  if ( $data->{'level'} eq '3' && $data->{'invoice'} ne '' ) { 
    if ( $old_invoice ne $data->{'invoice'} ) {
      $shpt_facture = $shpt_facture . $data->{'invoice'} . ';'; 
      $old_invoice = $data->{'invoice'};
    }
    $f_CURRENCY = $data->{'currency'};
    $f_RATE = $data->{'rate'};
    if ( $data->{'cost_kind'} eq 'FREIGHT' ) { 
      $f_COST_FREIGHT = $f_COST_FREIGHT + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
    }
    if ( $data->{'cost_kind'} eq 'PORT SERVICES' ) { 
      $f_COST_PORT = $f_COST_PORT + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
    }
    if ( $data->{'cost_kind'} eq 'CUSTOMS' ) { 
      $f_COST_CUSTOM = $f_COST_CUSTOM + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
    }
    # Modifié le 30/11/2010 : si on a des factures de saisie, 
    # on prend en compte les 2 postes pour le facturé
    #if ( substr($data->{'dpt'},0,4) eq 'Ind_' ) {
      if ( $data->{'cost_kind'} eq 'FINAL DELIVERY' ) { 
        $f_COST_FINAL = $f_COST_FINAL + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
      }
      if ( $data->{'cost_kind'} eq 'WAREHOUSING' ) { 
	$f_COST_WHSE = $f_COST_WHSE + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
      }
    #}
	
	
	if ( $data->{'cost_kind'} eq 'PURCHASE FEES' ) { 
      $f_PURCHASE_FEES = $f_PURCHASE_FEES + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
    }
	
    if ( $data->{'cost_kind'} eq 'ADMINISTRATIVE LOADS' ) { 
      $f_COST_ADMIN = $f_COST_ADMIN + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
    }
    if ( $data->{'cost_kind'} eq 'OTHER COST' ) { 
      $f_COST_OTHER = $f_COST_OTHER + $data->{'cost'}; 
      $f_total_cost = $f_total_cost + $data->{'cost'};
    }
  }

  # L'ENGAGE
  if ( $data->{'level'} eq '3' && $data->{'invoice'} eq '' ) { 
##### Modifié le 15 Juin 2012 : pour gérer le cas de POs avec même po_root dans un même BL
    if ( $old_bl ne $data->{'bl'} || $old_po ne $data->{'po'}  ) {
        $pcs_engage = $pcs_engage + $data->{'pcs'}; 
        $cbm_engage = $cbm_engage + $data->{'cbm'}; 
    }
#####
    if ( $old_bl ne $data->{'bl'} ) {
      $shpt_bl = $shpt_bl . $data->{'bl'} . ';'; 
##### Modifié le 15 Juin 2012 : pour gérer le cas de POs avec même po_root dans un même BL
#      if ( $data->{'date_arrived_pod'} != '0' ) {
#        $pcs_engage = $pcs_engage + $data->{'pcs'}; 
#        $cbm_engage = $cbm_engage + $data->{'cbm'}; 
#      }
#####
      $old_bl = $data->{'bl'};
    }
# Supprimé le 16 Novembre 2012, Demande Laetitia Durand, on prend le flottant donc dés que dans le DAU
#    if ( $data->{'date_arrived_pod'} != '0' ) {
      $e_CURRENCY = $data->{'currency'};
      $e_RATE = $data->{'rate'};
      if ( $data->{'cost_kind'} eq 'FREIGHT' ) { 
	$e_COST_FREIGHT = $e_COST_FREIGHT + $data->{'cost'}; 
        $e_total_cost = $e_total_cost + $data->{'cost'};
      }
      if ( $data->{'cost_kind'} eq 'PORT SERVICES' ) { 
	$e_COST_PORT = $e_COST_PORT + $data->{'cost'}; 
        $e_total_cost = $e_total_cost + $data->{'cost'};
      }
      if ( $data->{'cost_kind'} eq 'CUSTOMS' ) { 
	$e_COST_CUSTOM = $e_COST_CUSTOM + $data->{'cost'}; 
        $e_total_cost = $e_total_cost + $data->{'cost'};
      }
      if ( substr($data->{'dpt'},0,4) eq 'Ind_' ) {
        if ( $data->{'cost_kind'} eq 'FINAL DELIVERY' ) { 
	  $e_COST_FINAL = $e_COST_FINAL + $data->{'cost'}; 
          $e_total_cost = $e_total_cost + $data->{'cost'};
	}
        if ( $data->{'cost_kind'} eq 'WAREHOUSING' ) { 
	  $e_COST_WHSE = $e_COST_WHSE + $data->{'cost'}; 
          $e_total_cost = $e_total_cost + $data->{'cost'};
	}
      }
	  
	  if ( $data->{'cost_kind'} eq 'PURCHASE FEES' ) { 
	    $e_PURCHASE_FEES = $e_PURCHASE_FEES + $data->{'cost'}; 
        $e_total_cost = $e_total_cost + $data->{'cost'};
      }
	  
      if ( $data->{'cost_kind'} eq 'ADMINISTRATIVE LOADS' ) { 
	$e_COST_ADMIN = $e_COST_ADMIN + $data->{'cost'}; 
        $e_total_cost = $e_total_cost + $data->{'cost'};
      }
      if ( $data->{'cost_kind'} eq 'OTHER COST' ) { 
	$e_COST_OTHER = $e_COST_OTHER + $data->{'cost'}; 
        $e_total_cost = $e_total_cost + $data->{'cost'};
      }
#    }
  }

  if ( $data->{'level'} eq '3' ) { 
    # RECHERCHE des taux fees
    if ( $old_po  ne $data->{'po'} ) {
	open( LOG, '>> /home/fcs/clients/yrocher/bin/dau_to_excel_cd.log' );

		$rqlr_fees->execute( $data->{'po'} );
		if ( $dbh->errstr ne undef ) {
			print LOG "\n";
		   print LOG  $dbh->errstr.":\n"."num_po".$data->{'po'}."\n".$sqlr_fees."\n"; 
			print LOG "\n";		
		   exit;
		}
	close LOG;

      my $tx_import_fees = 0;
      my $tx_qc_fees = 0;
      my $tx_import_risk_fees = 0;
      my $tx_purchase_fees = 0;
      my $tx_other_fees = 0;
      while (  my $fees = $rqlr_fees->fetchrow_hashref ) {
        $tx_import_fees = $fees->{'import_fees_rate'};
        $tx_qc_fees = $fees->{'qc_fees_rate'};
        $tx_import_risk_fees = $fees->{'import_risk_fees_rate'};
        $tx_purchase_fees = $fees->{'purchase_fees_rate'};
        $tx_other_fees = $fees->{'other_fees_rate'};
      }
      $rqlr_fees->finish;
        
      # POUR LES COLONNES FACTUREES
      $rfeesfacture->execute( $data->{'po'} );
      while ( my $tabce = $rfeesfacture->fetchrow_hashref ) {
        $f_COST_FEES = $f_COST_FEES + ( $tabce->{'pcs'} * (($tabce->{'prix_fob2'}*$tx_import_fees) + ($tabce->{'prix_fob2'}*$tx_qc_fees) + ($tabce->{'prix_fob2'}*$tx_import_risk_fees) + ($tabce->{'prix_fob2'}*$tx_purchase_fees) + ($tabce->{'prix_fob2'}*$tx_other_fees) ));
      }
      $rfeesfacture->finish;

      # POUR LES COLONNES ENGAGEES
      $rfeesengage->execute( $data->{'po'} );
      while ( my $tabce = $rfeesengage->fetchrow_hashref ) {
        $e_COST_FEES = $e_COST_FEES + ( $tabce->{'pcs'} * (($tabce->{'prix_fob2'}*$tx_import_fees) + ($tabce->{'prix_fob2'}*$tx_qc_fees) + ($tabce->{'prix_fob2'}*$tx_import_risk_fees) + ($tabce->{'prix_fob2'}*$tx_purchase_fees) + ($tabce->{'prix_fob2'}*$tx_other_fees) ));
      }
      $rfeesengage->finish;
    }
    
    # Ajouter Octobre 2012 : Ecarts taux de couverture engagé et facturé
    if ( $data->{'level'} eq '3' ) { 
      if ( $old_po ne $data->{'po'} ) {
#        $recart_engage->execute(  $data->{'po'}, $data->{'po'}, $data->{'po'}, $data->{'po'}, $data->{'po'}, $data->{'po'} );
        my $cout_reel_engage = '';
        my $cout_corrige_engage = '';
        my $cout_reel_facture = '';
        my $cout_corrige_facture = '';
#        while ( my $resultat = $recart_engage->fetchrow_hashref ) {
#          $cout_reel_engage = $resultat->{'cout_reel_engage'};
#          $cout_corrige_engage = $resultat->{'cout_corrige_engage'};
#          $cout_reel_facture = $resultat->{'cout_reel_facture'};
#          $cout_corrige_facture = $resultat->{'cout_corrigé_facture'};
#          $ecart_engage += $cout_reel_engage - $cout_corrige_engage;
#          $ecart_facture += $cout_reel_facture - $cout_corrige_facture;
#        }
        $recart_engage->finish;
      }
    }
    $old_po =  $data->{'po'};
  }
}

# ON ECRIT LA DERNIERE LIGNE
############################
my %hash = ();
    $hash{UO} = $UO;
    $hash{STATUT} = $STATUT;
    $hash{CONTACT} = $CONTACT;
    if ( $FACTAVENIR eq '0' ) { $FACTAVENIR = 'Oui'; }
    else { $FACTAVENIR = 'Non'; }
    $hash{FACTAVENIR} = $FACTAVENIR;
    $hash{MOISOUVERT} = substr($MOISOUVERT,0,4).'-'.substr($MOISOUVERT,4,2);
    $hash{MOISSOLDE} = substr($MOISSOLDE,0,4).'-'.substr($MOISSOLDE,4,2);
    $hash{MOISCONFIRM} = substr($MOISCONFIRM,0,4).'-'.substr($MOISCONFIRM,4,2);
    $hash{MOISCLOTURE} = substr($MOISCLOTURE,0,4).'-'.substr($MOISCLOTURE,4,2);
    $hash{PO_ROOT} = $PO_ROOT;
    $hash{SHPT_COT} = $shpt_cot;
    $hash{SHPT_AMENDE} = $shpt_amende;
    $hash{SHPT_BL} = $shpt_bl;
    $hash{SHPT_FACTURE} = $shpt_facture;
	$hash{PURCHASE_RATE} = $purchase_fees;
    $hash{C_FREIGHT} = sprintf ("%0.1f", $c_COST_FREIGHT);
    $hash{C_PORT} = sprintf ("%0.1f", $c_COST_PORT);
    $hash{C_CUSTOM} = sprintf ("%0.1f", $c_COST_CUSTOM);
    $hash{C_FINAL} = sprintf ("%0.1f", $c_COST_FINAL);
    $hash{C_WHSE} = sprintf ("%0.1f", $c_COST_WHSE);
    $hash{C_FEES} = sprintf ("%0.1f", $c_COST_FEES);
	$hash{C_PURCHASE_FEES} = sprintf ("%0.1f", $c_PURCHASE_FEES);
    $hash{C_ADMIN} = sprintf ("%0.1f", $c_COST_ADMIN);
    $hash{C_OTHER} = sprintf ("%0.1f", $c_COST_OTHER);
    $hash{C_TOTAL} = sprintf ("%0.1f", $c_total_cost);
    $hash{C_CURRENCY} = $c_CURRENCY;
    $hash{C_RATE} = $c_RATE;
	
	#Recherche des containers

	$hash{C_LCL} = '';
	$hash{C_20} = '';
	$hash{C_20FR} = '';
	$hash{C_40} = '';
	$hash{C_40FR} = '';
	$hash{C_45} = '';
	$hash{C_40HC} = '';
	$hash{C_40HR} = '';
	$hash{C_CFS} = '';
	$hash{C_AIR} = '';


   $rcontainer_fcs->execute('1',$UO);

	while (my $container = $rcontainer_fcs->fetchrow_hashref) {

	  $hash{C_LCL} = $container->{'nb_lcl'}; 
	  $hash{C_20} = $container->{'nb_20'}; 
	  $hash{C_40} = $container->{'nb_40'}; 
	  $hash{C_45} =  $container->{'nb_45'};  
	  $hash{C_40HC} = $container->{'nb_40hc'}; 
	  $hash{C_20FR} = $container->{'nb_20fr'}; 
	  $hash{C_40FR} = $container->{'nb_40fr'};
	  $hash{C_40HR} = $container->{'nb_40hr'};
	  $hash{C_CFS} = $container->{'nb_cfs'};
	  $hash{C_AIR} = $container->{'nb_air'};
    }

	$rcontainer_fcs->finish;


    $hash{A_FREIGHT} = sprintf ("%0.1f", $a_COST_FREIGHT);
    $hash{A_PORT} = sprintf ("%0.1f", $a_COST_PORT);
    $hash{A_CUSTOM} = sprintf ("%0.1f", $a_COST_CUSTOM);
    $hash{A_FINAL} = sprintf ("%0.1f", $a_COST_FINAL);
    $hash{A_WHSE} = sprintf ("%0.1f", $a_COST_WHSE);
    $hash{A_FEES} = sprintf ("%0.1f", $a_COST_FEES);
	$hash{A_PURCHASE_FEES} = sprintf ("%0.1f", $a_PURCHASE_FEES);
    $hash{A_ADMIN} = sprintf ("%0.1f", $a_COST_ADMIN);
    $hash{A_OTHER} = sprintf ("%0.1f", $a_COST_OTHER);
    $hash{A_TOTAL} = sprintf ("%0.1f", $a_total_cost);
    $hash{A_CURRENCY} = $a_CURRENCY;
    $hash{A_RATE} = $a_RATE;


# Recherche des containers

	$hash{A_LCL} = '';
	$hash{A_20} = '';
	$hash{A_20FR} = '';
	$hash{A_40} = '';
	$hash{A_40FR} = '';
	$hash{A_45} = '';
	$hash{A_40HC} = '';
	$hash{A_40HR} = '';
	$hash{A_CFS} = '';
	$hash{A_AIR} = '';


$rcontainer_fcs->execute('2',$UO);

	while (my $container = $rcontainer_fcs->fetchrow_hashref) {

	  $hash{A_LCL} = $container->{'nb_lcl'}; 
	  $hash{A_20} = $container->{'nb_20'}; 
	  $hash{A_40} = $container->{'nb_40'}; 
	  $hash{A_45} =  $container->{'nb_45'};  
	  $hash{A_40HC} = $container->{'nb_40hc'}; 
	  $hash{A_20FR} = $container->{'nb_20fr'}; 
	  $hash{A_40FR} = $container->{'nb_40fr'};
	  $hash{A_40HR} = $container->{'nb_40hr'};
	  $hash{A_CFS} = $container->{'nb_cfs'};
	  $hash{A_AIR} = $container->{'nb_air'};
    }

	$rcontainer_fcs->finish;

    $hash{F_FREIGHT} = sprintf ("%0.1f", $f_COST_FREIGHT);
    $hash{F_PORT} = sprintf ("%0.1f", $f_COST_PORT);
    $hash{F_CUSTOM} = sprintf ("%0.1f", $f_COST_CUSTOM);
    $hash{F_FINAL} = sprintf ("%0.1f", $f_COST_FINAL);
    $hash{F_WHSE} = sprintf ("%0.1f", $f_COST_WHSE);
    $hash{F_FEES} = sprintf ("%0.1f", $f_COST_FEES);
	$hash{F_PURCHASE_FEES} = sprintf ("%0.1f", $f_PURCHASE_FEES);
    $hash{F_ADMIN} = sprintf ("%0.1f", $f_COST_ADMIN);
    $hash{F_OTHER} = sprintf ("%0.1f", $f_COST_OTHER);
    $f_total_cost = $f_total_cost + $f_COST_FEES;
    $hash{F_TOTAL} = sprintf ("%0.1f", $f_total_cost);
    $hash{F_CURRENCY} = $f_CURRENCY;
    $hash{F_RATE} = $f_RATE;
    if ( $FACTAVENIR eq 'Oui' ) {
      $hash{E_FREIGHT} = sprintf ("%0.1f", $e_COST_FREIGHT);
      $hash{E_PORT} = sprintf ("%0.1f", $e_COST_PORT);
      $hash{E_CUSTOM} = sprintf ("%0.1f", $e_COST_CUSTOM);
      $hash{E_FINAL} = sprintf ("%0.1f", $e_COST_FINAL);
      $hash{E_WHSE} = sprintf ("%0.1f", $e_COST_WHSE);
      $hash{E_FEES} = sprintf ("%0.1f", $e_COST_FEES);
	  $hash{E_PURCHASE_FEES} = sprintf ("%0.1f", $e_PURCHASE_FEES);
      $hash{E_ADMIN} = sprintf ("%0.1f", $e_COST_ADMIN);
      $hash{E_OTHER} = sprintf ("%0.1f", $e_COST_OTHER);
      $e_total_cost = $e_total_cost + $e_COST_FEES;
      $hash{E_TOTAL} = sprintf ("%0.1f", $e_total_cost);
      $hash{E_CURRENCY} = $e_CURRENCY;
      $hash{E_RATE} = $e_RATE;
    }
    else {
      $hash{E_FREIGHT} = sprintf ("%0.1f", $f_COST_FREIGHT);
      $hash{E_PORT} = sprintf ("%0.1f", $f_COST_PORT);
      $hash{E_CUSTOM} = sprintf ("%0.1f", $f_COST_CUSTOM);
      $hash{E_FINAL} = sprintf ("%0.1f", $f_COST_FINAL);
      $hash{E_WHSE} = sprintf ("%0.1f", $f_COST_WHSE);
      $hash{E_FEES} = sprintf ("%0.1f", $f_COST_FEES);
	  $hash{E_PURCHASE_FEES} = sprintf ("%0.1f", $f_PURCHASE_FEES);
      $hash{E_ADMIN} = sprintf ("%0.1f", $f_COST_ADMIN);
      $hash{E_OTHER} = sprintf ("%0.1f", $f_COST_OTHER);
      $hash{E_TOTAL} = sprintf ("%0.1f", $f_total_cost);
      $hash{E_CURRENCY} = $f_CURRENCY;
      $hash{E_RATE} = $f_RATE;
    }
	

	# Recherche des containers
	$hash{E_LCL} = '';
	$hash{E_20} = '';
	$hash{E_20FR} = '';
	$hash{E_40} = '';
	$hash{E_40FR} = '';
	$hash{E_45} = '';
	$hash{E_40HC} = '';
	$hash{E_40HR} = '';
	$hash{E_CFS} = '';
	$hash{E_AIR} = '';

	$rcontainer_fcs->execute('3',$UO);

	while (my $container = $rcontainer_fcs->fetchrow_hashref) {

	  $hash{E_LCL} = $container->{'nb_lcl'}; 
	  $hash{E_20} = $container->{'nb_20'}; 
	  $hash{E_40} = $container->{'nb_40'}; 
	  $hash{E_45} =  $container->{'nb_45'};  
	  $hash{E_40HC} = $container->{'nb_40hc'}; 
	  $hash{E_20FR} = $container->{'nb_20fr'}; 
	  $hash{E_40FR} = $container->{'nb_40fr'};
	  $hash{E_40HR} = $container->{'nb_40hr'};
	  $hash{E_CFS} = $container->{'nb_cfs'};
	  $hash{E_AIR} = $container->{'nb_air'};
    }

	$rcontainer_fcs->finish;

    $hash{ECART_AC_FREIGHT} = sprintf ("%0.1f", $a_COST_FREIGHT - $c_COST_FREIGHT);
    $hash{ECART_AC_PORT} = sprintf ("%0.1f", $a_COST_PORT - $c_COST_PORT);
    $hash{ECART_AC_CUSTOM} = sprintf ("%0.1f", $a_COST_CUSTOM - $c_COST_CUSTOM);
    $hash{ECART_AC_FINAL} = sprintf ("%0.1f", $a_COST_FINAL - $c_COST_FINAL);
    $hash{ECART_AC_WHSE} = sprintf ("%0.1f", $a_COST_WHSE - $c_COST_WHSE);
    $hash{ECART_AC_FEES} = sprintf ("%0.1f", $a_COST_FEES - $c_COST_FEES);
    $hash{ECART_AC_ADMIN} = sprintf ("%0.1f", $a_COST_ADMIN - $c_COST_ADMIN);
    $hash{ECART_AC_OTHER} = sprintf ("%0.1f", $a_COST_OTHER - $c_COST_OTHER);
    $hash{TOTAL_ECART_AC} = sprintf ("%0.1f", $a_total_cost - $c_total_cost);
    $hash{ECART_AF_FREIGHT} = sprintf ("%0.1f", $f_COST_FREIGHT - $c_COST_FREIGHT);
    $hash{ECART_AF_PORT} = sprintf ("%0.1f", $f_COST_PORT - $c_COST_PORT);
    $hash{ECART_AF_CUSTOM} = sprintf ("%0.1f", $f_COST_CUSTOM - $c_COST_CUSTOM);
    $hash{ECART_AF_FINAL} = sprintf ("%0.1f", $f_COST_FINAL - $c_COST_FINAL);
    $hash{ECART_AF_WHSE} = sprintf ("%0.1f", $f_COST_WHSE - $c_COST_WHSE);
    $hash{ECART_AF_FEES} = sprintf ("%0.1f", $f_COST_FEES - $c_COST_FEES);
    $hash{ECART_AF_ADMIN} = sprintf ("%0.1f", $f_COST_ADMIN - $c_COST_ADMIN);
    $hash{ECART_AF_OTHER} = sprintf ("%0.1f", $f_COST_OTHER - $c_COST_OTHER);
    $hash{TOTAL_ECART_AF} = sprintf ("%0.1f", $f_total_cost - $c_total_cost);
    if ( $FACTAVENIR eq 'Oui' ) {
      $hash{ECART_EF_FREIGHT} = sprintf ("%0.1f", $e_COST_FREIGHT - $f_COST_FREIGHT);
      $hash{ECART_EF_PORT} = sprintf ("%0.1f", $e_COST_PORT - $f_COST_PORT);
      $hash{ECART_EF_CUSTOM} = sprintf ("%0.1f", $e_COST_CUSTOM - $f_COST_CUSTOM);
      $hash{ECART_EF_FINAL} = sprintf ("%0.1f", $e_COST_FINAL - $f_COST_FINAL);
      $hash{ECART_EF_WHSE} = sprintf ("%0.1f", $e_COST_WHSE - $f_COST_WHSE);
      $hash{ECART_EF_FEES} = sprintf ("%0.1f", $e_COST_FEES - $f_COST_FEES);
      $hash{ECART_EF_ADMIN} = sprintf ("%0.1f", $e_COST_ADMIN - $f_COST_ADMIN);
      $hash{ECART_EF_OTHER} = sprintf ("%0.1f", $e_COST_OTHER - $f_COST_OTHER);
      $hash{TOTAL_ECART_EF} = sprintf ("%0.1f", $e_total_cost - $f_total_cost);
    }
    else {
      $hash{ECART_EF_FREIGHT} = '0' ;
      $hash{ECART_EF_PORT} = '0';
      $hash{ECART_EF_CUSTOM} = '0';
      $hash{ECART_EF_FINAL} = '0';
      $hash{ECART_EF_WHSE} = '0';
      $hash{ECART_EF_FEES} = '0';
      $hash{ECART_EF_ADMIN} = '0';
      $hash{ECART_EF_OTHER} = '0';
      $hash{TOTAL_ECART_EF} = '0';
    }
    if ( $pcs_cot == 0 ) { $pcs_cot = 1; }
    my $bis_freight = $c_COST_FREIGHT / $pcs_cot * $pcs_engage;
    my $bis_port = $c_COST_PORT / $pcs_cot * $pcs_engage;
    my $bis_custom = $c_COST_CUSTOM / $pcs_cot * $pcs_engage;
    my $bis_final = $c_COST_FINAL / $pcs_cot * $pcs_engage;
    my $bis_whse = $c_COST_WHSE / $pcs_cot * $pcs_engage;
    my $bis_fees = $c_COST_FEES / $pcs_cot * $pcs_engage;
    my $bis_admin = $c_COST_ADMIN / $pcs_cot * $pcs_engage;
    my $bis_other = $c_COST_OTHER / $pcs_cot * $pcs_engage;
    my $bis_total = $c_total_cost / $pcs_cot * $pcs_engage;
    if ( $FACTAVENIR eq 'Oui' ) {
      $hash{ECART_BE_FREIGHT} = sprintf ("%0.1f", $bis_freight - $e_COST_FREIGHT);
      $hash{ECART_BE_PORT} = sprintf ("%0.1f", $bis_port - $e_COST_PORT);
      $hash{ECART_BE_CUSTOM} = sprintf ("%0.1f", $bis_custom - $e_COST_CUSTOM);
      $hash{ECART_BE_FINAL} = sprintf ("%0.1f", $bis_final - $e_COST_FINAL);
      $hash{ECART_BE_WHSE} = sprintf ("%0.1f", $bis_whse - $e_COST_WHSE);
      $hash{ECART_BE_FEES} = sprintf ("%0.1f", $bis_fees - $e_COST_FEES);
      $hash{ECART_BE_ADMIN} = sprintf ("%0.1f", $bis_admin - $e_COST_ADMIN);
      $hash{ECART_BE_OTHER} = sprintf ("%0.1f", $bis_other - $e_COST_OTHER);
      $hash{TOTAL_ECART_BE} = sprintf ("%0.1f", $bis_total - $e_total_cost);
    }
    else {
      $hash{ECART_BE_FREIGHT} = sprintf ("%0.1f", $bis_freight - $f_COST_FREIGHT);
      $hash{ECART_BE_PORT} = sprintf ("%0.1f", $bis_port - $f_COST_PORT);
      $hash{ECART_BE_CUSTOM} = sprintf ("%0.1f", $bis_custom - $f_COST_CUSTOM);
      $hash{ECART_BE_FINAL} = sprintf ("%0.1f", $bis_final - $f_COST_FINAL);
      $hash{ECART_BE_WHSE} = sprintf ("%0.1f", $bis_whse - $f_COST_WHSE);
      $hash{ECART_BE_FEES} = sprintf ("%0.1f", $bis_fees - $f_COST_FEES);
      $hash{ECART_BE_ADMIN} = sprintf ("%0.1f", $bis_admin - $f_COST_ADMIN);
      $hash{ECART_BE_OTHER} = sprintf ("%0.1f", $bis_other - $f_COST_OTHER);
      $hash{TOTAL_ECART_BE} = sprintf ("%0.1f", $bis_total - $f_total_cost);
    }
    $hash{PCS_COT} = $pcs_cot;
    $hash{PCS_AMENDE} = $pcs_amende;
    $hash{PCS_ENGAGE} = $pcs_engage;
    $hash{CBM_COT} = sprintf("%0.2f",$cbm_cot);
    $hash{CBM_AMENDE} = sprintf("%0.2f",$cbm_amende);
    $hash{CBM_ENGAGE} = sprintf("%0.2f",$cbm_engage);
    
    # CA Import Net
    my $c_caimportnet = 0;
    my $a_caimportnet = 0;
    $rCAimport->execute( $PO_ROOT );
    while ( my $tabca = $rCAimport->fetchrow_hashref ) {
      $c_caimportnet = $c_caimportnet + ( $tabca->{'pcs'} * ($tabca->{'prix_hub'} - $tabca->{'prix_fob2'}) ); 
      $rphf->execute( $PO_ROOT, '2', $tabca->{'sku'},  $tabca->{'size'},  $tabca->{'color'}); 
      while ( my $tabsku = $rphf->fetchrow_hashref) {
        $a_caimportnet = $a_caimportnet + ( $tabsku->{'pcs'} * ($tabca->{'prix_hub'} - $tabca->{'prix_fob2'}) );
      }
      $rphf->finish;
    }
    $rCAimport->finish;
    $hash{C_CAIMPORTNET} = sprintf ("%0.1f", $c_caimportnet - $c_COST_FEES);
    $hash{A_CAIMPORTNET} = sprintf ("%0.1f", $a_caimportnet - $a_COST_FEES);

    # POUR l'engagé
    ################
    my $e_caimportnet = 0;
    $rCAimportengage->execute( $PO_ROOT );
    while ( my $tabce = $rCAimportengage->fetchrow_hashref ) {
      $e_caimportnet = $e_caimportnet + ( $tabce->{'pcs'} * ($tabce->{'prix_facture'} - $tabce->{'prix_fob2'}) ); 
    }
    $rCAimportengage->finish;
    $hash{E_CAIMPORTNET} = sprintf ("%0.1f", $e_caimportnet - $e_COST_FEES);
    
    # POUR le facturé
    ################
    my $f_caimportnet = 0;
    $rCAimportfacture->execute( $PO_ROOT );
    while ( my $tabce = $rCAimportfacture->fetchrow_hashref ) {
      $f_caimportnet = $f_caimportnet + ( $tabce->{'pcs'} * ($tabce->{'prix_facture'} - $tabce->{'prix_fob2'}) ); 
    }
    $rCAimportfacture->finish;
    $hash{F_CAIMPORTNET} = sprintf ("%0.1f", $f_caimportnet - $f_COST_FEES);

    # On recherche les surcouts 
    $rc1->execute( $PO_ROOT );
    my $surcout_ouvert = $rc1->fetchrow;
    $rc1->finish;
    $hash{SURCOUTOUVERT} = $surcout_ouvert;
    $rc2->execute( $PO_ROOT );
    my $surcout_rapproche = $rc2->fetchrow;
    $rc2->finish;
    $hash{SURCOUTRAPPROCHE} = $surcout_rapproche;
    $rc3->execute( $PO_ROOT );
    my $surcout_confirme = $rc3->fetchrow;
    $rc3->finish;
    $hash{SURCOUTCONFIRME} = $surcout_confirme;
    $rc4->execute( $PO_ROOT );
    my $surcout_cloture = $rc4->fetchrow;
    $rc4->finish;
    $hash{SURCOUTCLOTURE} = $surcout_cloture;
    $hash{TOTALSURCOUT} = $surcout_ouvert + $surcout_rapproche + $surcout_confirme + $surcout_cloture ;
    
    my $commentsurcout = '';
    $rdec->execute( $PO_ROOT );
    while ( my @result = $rdec->fetchrow ) {
      $commentsurcout = $commentsurcout . $result[0] . ' ';
    }
    $rdec->finish;
    $hash{COMMENTSURCOUT} = $commentsurcout;

    $hash{COMMENTCA} = '';

    my $commentop = '';
    $rdu->execute( $UO );
    while ( my @result = $rdu->fetchrow ) {
      $commentop = $commentop . $result[0] . ' ';
    }
    $rdu->finish;
    $hash{COMMENTOP} = $commentop;

    my $commentporoot = '';
    $rdpr->execute( $PO_ROOT );
    while ( my @result = $rdpr->fetchrow ) {
      $commentporoot = $commentporoot . $result[0] . ' ';
    }
    $rdpr->finish;
    $hash{COMMENTPOROOT} = $commentporoot;

    my $commentpo = '';
    $rdp->execute( $PO );
    while ( my @result = $rdp->fetchrow ) {
      $commentpo = $commentpo . $result[0] . ' ';
    }
    $rdp->finish;
    $hash{COMMENTPO} = $commentpo;

    #$hash{C_RESULTATOP} = $hash{C_CAIMPORTNET} - $hash{C_TOTAL};
    #$hash{A_RESULTATOP} = $hash{A_CAIMPORTNET} - $hash{A_TOTAL};
    #$hash{F_RESULTATOP} = $hash{F_CAIMPORTNET} - $hash{F_TOTAL} - $hash{SURCOUTOUVERT};
    #$hash{E_RESULTATOP} = $hash{E_CAIMPORTNET} - $hash{E_TOTAL};
    # Les Fees sont déjà déduites dans le CA Import Net
    $hash{C_RESULTATOP} = $hash{C_CAIMPORTNET} - $hash{C_TOTAL} + $hash{C_FEES};
    $hash{A_RESULTATOP} = $hash{A_CAIMPORTNET} - $hash{A_TOTAL} + $hash{A_FEES};
    $hash{F_RESULTATOP} = $hash{F_CAIMPORTNET} - $hash{F_TOTAL} + $hash{F_FEES} - $hash{SURCOUTOUVERT};
    $hash{E_RESULTATOP} = $hash{E_CAIMPORTNET} - $hash{E_TOTAL} + $hash{E_FEES};
    $hash{ECART_ENGAGE} = sprintf ("%0.1f",$ecart_engage);
    $hash{ECART_FACTURE} = sprintf ("%0.1f",$ecart_facture);

push ( @loop, \%hash );

$rq->finish;

$template->param( loop    => \@loop );


$dbh->disconnect;
print $template->output;
if ( $DBUG eq '1' ) { 
  my $date_fin = now;
  print FICHIER_LOG "\n\nFIN $date_fin";
  close FICHIER_LOG; 
}




sub init(){
# On définit les options obligatoires et l'attribution des options aux variables globales;
	$log_msg .= "Debut du programme :".`date`."\n";
	$log_msg .= "Liste des  arguments: $0 @ARGV \n";

	for ( my $i ; $i < scalar @ARGV ; $i++ ){
		if (get_arg($ARGV[$i]) eq 'o'){
			$is_original = $ARGV[$i];
		} elsif (get_arg($ARGV[$i]) eq 'c'){
			$is_checking = $ARGV[$i];
		} elsif (get_arg($ARGV[$i]) eq 'e'){
			$is_embarque = 1;
		} elsif (get_arg($ARGV[$i]) eq 'f'){
			$is_facture = 1;
		}elsif (get_arg($ARGV[$i]) eq 'u' ){
			$uo_num = $ARGV[$i+1];
		}elsif (get_arg($ARGV[$i]) eq 'p' ){
			$num_po = $ARGV[$i+1];
		}
	}
}



sub is_elt_exist_in_array {
	my ($elt_to_test, @array  ) = @_;
	$value_to_return = 0;
	foreach my $cur_elt (@array) {
		if ($cur_elt eq $elt_to_test) {
			$value_to_return = 1
		}
	}
	return $value_to_return;

}


sub init_error(){
# On décrit les options de l'application;
  $log_msg .= "usage : dau_to_excel_generate.pl -o|-c|-e|-f\n";
  &exit_function;
}

sub get_arg(){
        my ($str_to_return) = @_;
        if (index( $str_to_return, '-' ) > -1){
                $str_to_return = substr($str_to_return , 1, 1);
        } else {
                $str_to_return = 0
        }
        return ($str_to_return);
}

sub exit_function(){
	if ($is_print_into_log){
		my $fichier_log = substr($0, 0, rindex ($0, '.')).".log";
		if(!-e $fichier_log) {
			`touch $fichier_log`;
			`chmod 777 $fichier_log`;
		}
		open( LOG, ">> $fichier_log" ) || die "Impossible d'ouvrir fichier log !\n";
		
		print LOG $log_msg;
		close LOG;
	}
	print $log_msg;
   exit;
}



1694