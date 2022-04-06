<?php require 'vendor/autoload.php';
use Symfony\Component\DomCrawler\Crawler;
use Symfony\Component\CssSelector\CssSelector;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\CssSelector\CssSelectorConverter;

ini_set('max_execution_time', 0);
ini_set('memory_limit', '-1');

$client = new \GuzzleHttp\Client(['verify' => false]);
// $base_url = 'https://haisha-yoyaku.jp/bun2sdental/list/?page=1';
$start_row = 2; // first row for column headers

if(isset($_GET['end'])){
    $end = $_GET['end'];
}else{
    $end = 1;
}

if(isset($_GET['start'])){
    $start = $_GET['start'];
    if ($end == 1){
        $end = $start;
    }
}else{
    $start = 1;
}

$batch = [];
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("epark.xlsx"); // read existing spreadsheet
// $spreadsheet = new Spreadsheet;
$sheet = $spreadsheet->getActiveSheet();
$count = $start_row;
    $sheet->setCellValue('A1', 'URL');
    $sheet->setCellValue('B1', 'Clinic Name');
    $sheet->setCellValue('C1', 'Furigana');
    $sheet->setCellValue('D1', 'Phone Number');
    $sheet->setCellValue('E1', 'Address');
    $sheet->setCellValue('F1', 'Closed Days');
    $sheet->setCellValue('G1', 'Remark1'); 
    $sheet->setCellValue('H1', 'Remark2');
    $sheet->setCellValue('I1', 'Medical Items');
    $sheet->setCellValue('J1', 'Facility Information');
    $sheet->setCellValue('K1', 'Services');
    $sheet->setCellValue('L1', 'Stations');
    $sheet->setCellValue('M1', 'Reception Hours');
    $sheet->setCellValue('N1', 'Monday');
    $sheet->setCellValue('O1', 'Tuesday');
    $sheet->setCellValue('P1', 'Wednesday');
    $sheet->setCellValue('Q1', 'Thursday');
    $sheet->setCellValue('R1', 'Friday');
    $sheet->setCellValue('S1', 'Saturday');
    $sheet->setCellValue('T1', 'Sunday');
    $sheet->setCellValue('U1', 'Holiday');

    $writer = new Xlsx($spreadsheet);

    $cols = ['M','N','O','P','Q','R','S','T','U']; 
for ( $i = $start; $i <= $end; $i++ ){

    $urls = 'https://haisha-yoyaku.jp/bun2sdental/list/?page='.$i.''; 
    $res = $client->request('GET', $urls);
    $html = ''.$res->getBody();
    $crawler = new Crawler($html);

    $nodeValues = $crawler->filter('.list_search_casette_upper')->each(function (Crawler $node, $i) use(&$sheet, &$count, &$spreadsheet, &$cols, &$writer){

        $x_client = new \GuzzleHttp\Client();

        // url scraped from landing page
        $x_url = $node->filter('a')->attr('href');

        # for testing empty nodes
        // $x_url = "https://haisha-yoyaku.jp/bun2sdental/detail/index/id/2735502805/";
        $x_res = $x_client->request('GET', $x_url);
        $x_html = ''.$x_res->getBody();
        $x_crawler = new Crawler($x_html);

        #crawl data inside scraped links
        $address="";
        $address_counter = 0;
        $address_position = 0;
        $addressCrawler = $x_crawler->filter('#clinic_basic-information > .section_column1 > .table_clinic-base tr > th')->each(function (Crawler $addressItem, $i) use (&$x_crawler, &$count, &$sheet, &$spreadsheet, &$address, &$address_counter, &$address_position) {

            $address_label = $addressItem->text();
            if ($address_label == "住所"){
                $address_position = $address_counter;
            }
            $address_counter++;
        });

        $address = $x_crawler->filter('#clinic_basic-information > .section_column1 > .table_clinic-base tr > td')->eq($address_position)->count() ? $x_crawler->filter('#clinic_basic-information > .section_column1 > .table_clinic-base tr > td')->eq($address_position)->text() : '';

        $services = "";
        $medicalCrawler = $x_crawler->filter('.detail_top_subject_wrap > span')->each(function (Crawler $medicalItem, $i) use (&$x_crawler, &$medical_items) {
            $medical_items .= $medicalItem->filter('span')->count()
            ? " ".$medicalItem->filter('span')->text()." /" 
            : '';
        });

        $stations = "";
        $stationCrawler = $x_crawler->filter('.detail_basic_info_walk_icon')->each(function (Crawler $stationItem, $i) use (&$x_crawler, &$stations) {
            $stations .= $stationItem->filter('.detail_basic_info_walk_icon')->count()
            ? " ".$stationItem->filter('.detail_basic_info_walk_icon')->text()." / \n" 
            : '';
        });
        
        $medical_items = "";
        $facility_items = "";
        $fac_counter = 0;
        $fac_position = 0;
        $med_counter = 0;
        $med_position = 0;
        $facilityCrawler = $x_crawler->filter('.area_section-detail02:not(#clinic_basic-information) > .section_column1 > .table_clinic-base tr > th')->each(function (Crawler $facItem, $i) use (&$x_crawler, &$facility_items, &$services, &$fac_counter, &$fac_position, &$med_counter, &$med_position) {
            
            $labels = $facItem->text();
            if ($labels == "施設情報"){
                $fac_position = $fac_counter;
            }else{
                $fac_position = 20; // set to a value of position with nothing to fetch
            }

            if ($labels == "サービス"){
                $med_position = $med_counter;
            }else{
                $med_position = 20; // set to a value of position with nothing to fetch
            }

            $fac_counter++;
            $med_counter++;
            
        });

        $medical_items = $x_crawler->filter('.area_section-detail02:not(#clinic_basic-information) > .section_column1 > .table_clinic-base tr > td')->eq($med_position)->count() ? $x_crawler->filter('.area_section-detail02:not(#clinic_basic-information) > .section_column1 > .table_clinic-base tr > td')->eq($med_position)->text() : '';
        $facility_items = $x_crawler->filter('.area_section-detail02:not(#clinic_basic-information) > .section_column1 > .table_clinic-base tr > td')->eq($fac_position)->count() ? $x_crawler->filter('.area_section-detail02:not(#clinic_basic-information) > .section_column1 > .table_clinic-base tr > td')->eq($fac_position)->text() : '';

        $name = $x_crawler->filter('.main:not(.main_kana)')->count() ? $x_crawler->filter('.main:not(.main_kana)')->text() : '';
        $furigana = $x_crawler->filter('.main_kana')->count() ? $x_crawler->filter('.main_kana')->text() : '';
        $new_name = str_replace($furigana, '', $name);
        $phone_number = $x_crawler->filter('.infomation_telephone')->count() ? $x_crawler->filter('.infomation_telephone')->text() : '';
        $closed_days = $x_crawler->filter('.closed_day_icon')->count() ? $x_crawler->filter('.closed_day_icon')->text() : '';
        $remark1 = $x_crawler->filter('.holiday_treatment_icon')->count() ? $x_crawler->filter('.holiday_treatment_icon')->text() : '';
        $remark2 = $x_crawler->filter('.night_treatment_icon')->count() ? $x_crawler->filter('.night_treatment_icon')->text() : '';

        $sheet->setCellValue('A' .$count. '', $x_url);
        $sheet->setCellValue('B' .$count. '', $new_name);
        $sheet->setCellValue('C' .$count. '', $furigana);
        $sheet->setCellValue('D' .$count. '', $phone_number);
        $sheet->setCellValue('E' .$count. '', $address);
        $sheet->setCellValue('F' .$count. '', $closed_days);
        $sheet->setCellValue('G' .$count. '', $remark1);
        $sheet->setCellValue('H' .$count. '', $remark2);
        $sheet->setCellValue('I' .$count. '', $medical_items);
        $sheet->setCellValue('J' .$count. '', $facility_items);
        $sheet->setCellValue('K' .$count. '', $services);
        $sheet->setCellValue('L' .$count. '', $stations);

        echo $x_url."<br>";
        echo $new_name."<br>";
        // echo $furigana."<br>";
        // echo $phone_number."<br>";
        echo $address."<br>";
        // echo $closed_days."<br>";
        // echo $remark1."<br>";
        // echo $remark2."<br>";
        // echo $medical_items."<br>";
        // echo $facility_items."<br>";
        // echo $services."<br>";
        // echo $stations."<br>";

        $schedules = [];
        $time = 0;
        $scheduleCrawler = $x_crawler->filter('.a_t-t > .t_c-b tr:not(.top) > th, .a_t-t > .t_c-b tr:not(.top) > td')->each(function (Crawler $schedItem, $i) use (&$x_crawler, &$sheet, &$spreadsheet, &$schedules, &$cols, &$count, &$time) {

            // $dates = $schedItem->filter('td')->count() ? $schedItem->filter('td')->text() : '';
            $dates = $schedItem->text();
            echo $dates;
            $location = $cols[$time].$count;
            $sheet->setCellValue($location, $dates);
            $current_col = $cols[$time];

            if($current_col == $cols[count($cols) -1]){
                $count++;
                $time = 0;
            }else{
                $time++;
            }
        });
        $count++;
        echo "<br><br>";

            $writer->save('epark.xlsx');
            sleep(3);

    });

    echo "Page ". $i . ": Done.<br>";
    sleep(2);

}

?>