#!/usr/bin/perl

use strict;
use warnings;

use JSON::XS qw(encode_json decode_json);
use Getopt::Long;
use LWP::Simple qw(get);
use Term::ANSIColor;
use HTML::Strip;
use Excel::Writer::XLSX;
use POSIX qw(strftime ceil);
use Term::ProgressBar;
use File::Basename qw(fileparse);
use Archive::Zip qw(:ERROR_CODES);
use MIME::Lite;
use Data::Dumper;
use Switch;
use Encode qw(decode encode);
use Devel::Size qw(size total_size);

=begin comment
1) Чтобы начать использовать: ./parse_gis.pl --k your_key
2) Заполнение тестовыми данными нужно чтобы понять какое количество данных упрется в предел CPU и RAM
3) 1 компания = 1 items в JSON
=cut

=response_code comment
200 Успешный запрос
400 Ошибка в запросе
403 Доступ запрещен
404 По запросу ничего не найдено
500 Внутренняя ошибка сервера
503 Сервис временно не доступен
=cut

###############################################################################
#Constants
###############################################################################
#Отладочный режим, т.е показывать всю информацию
use constant DEBUG 		=> 0;

#0 - выполняем парсинг, 1 - заполняем тестовыми данными		
use constant TEST 		=> 0;

#Количество строк тестовых данных		
use constant COUNT_TEST 	=> 10000;

#0 - не сохранять, 1 - сохранять	
use constant SAVE_TO_EXCEL 	=> 1;

#0 - каждая рубрика отдельный файл, 1 - писать все в один файл
use constant EXCEL_FULL => 0;

#0 - не сохранять, 1 - сохранять		
use constant SAVE_JSON_TO_FILE 	=> 1;

#0 - архивировать, 1 - не архивировать		
use constant ARCHIVE_EXCEL 	=> 0;

#0 - архивировать, 1 - не архивировать	
use constant ARCHIVE_JSON 	=> 0;

#Отчет по email
#0 - не посылать, 1 - посылать
use constant REPORT_VIA_EMAIL 	=> 1;
use constant FROM => 'report\@your_domain';
use constant RECIPIENT => 'info\@your_domain';
use constant SUBJECT => 'Отчет о сборе данных';
use constant SMTP_SERVER => '127.0.0.1';
use constant EMAIL_DEBUG => 1;

#Путь для сохранения
use constant PATH_FOR_SAVING => '/home/parse_gis/data/';

#Максимально кол-во компании которое нужно вытащить
#Используется для отладки
#Если 0 то значит все рубрики
use constant MAX_RUBRIC => 0;

###############################################################################
#Global variables
###############################################################################

my $KEY;
my $GENERAL_URL = "https://catalog.api.2gis.ru/2.0/catalog/";

#Идентификатор города
#1 - Новосибирск, 5 - Кемерово
my $REGION_ID = 5;

my $CITY;
my $COLOR_NAME = 'red on_bright_green';
my $COLOR_DELIMITER = 'yellow';

my %MAIN_DATA;

#for summary information
my $COUNT_GENERAL_RUBRIC = 0;
my $COUNT_RUBRIC = 0;
my $COUNT_COMPANY = 0;
my $COUNT_COMPANY_EMAIL = 0;
my $COUNT_COMPANY_WEBSITE = 0;

my $COUNT_COLS;

###############################################################################
#									START
###############################################################################
&main();

sub parse_argv
{
    GetOptions ('k=s' => \$KEY);
}

sub do_debug
{

}

sub get_date
{
   return strftime '%Y%m%d', localtime;
}
sub get_name
{	
	#my $date = strftime '%Y%m%d', localtime;
	my $date = get_date;
	
	switch ($_[1])
	{
		case "json"	{ return $_[0] . '_city-' . $REGION_ID . '_date-' . "$date.json" }
		case "xlsx"	{ return $_[0] . '_city-' . $REGION_ID . '_date-' . "$date.xlsx" }
	}
}

sub do_test
{
    fill_hash_test(COUNT_TEST);

	my $file_json = get_name('2gis_TEST', 'json');
	my $file_xlsx = get_name('2gis_TEST', 'xlsx');
	
    save_json_to_file($file_json) if SAVE_JSON_TO_FILE;

    save_to_excel($file_xlsx) if SAVE_TO_EXCEL;

    exit;
}

sub get_data
{
    my $url = $GENERAL_URL . "rubric/list?region_id=$REGION_ID&fields=items.rubrics&key=$KEY";

    my $response = get($url);

    my $decoded = decode_json($response);

    printlnEx('*' x $COUNT_COLS, 'yellow on_magenta');

    println('Версия API: ' . $decoded->{'meta'}{'api_version'}, 'bold green');
    println('Дата: ' . $decoded->{'meta'}{'issue_date'}, 'bold green');
    println('Общее количество рубрик: ' . $decoded->{'result'}{'total'}, 'bold green');

    printlnEx("*" x $COUNT_COLS, 'yellow on_magenta');

    sleep(5);

    printlnEx('******************START**********************', 'red on_bright_yellow');

    binmode STDOUT, ":utf8";

    my @items = @{$decoded->{'result'}{'items'}};

    my $item_count = 1;

    #Идем по рубрикам
    foreach my $item (@items)
    {
	#Для тестирования чтобы не парсить весь сайт расскоментировать строку
	#if ($item_count > MAX_RUBRIC) { last; }
	
	println("#: $item_count", 'bold green');
	
	my $item_name = $item->{'name'};
	printlnEx("Name: " . $item->{'name'}, $COLOR_NAME);
	
	
	println("ID: " . $item->{'id'} . "\n");
	println("org_count: " . $item->{'org_count'});
	println("branch_count: " . $item->{'branch_count'});
	println("Type: " . $item->{'type'});

	#$item_count++;

	my $rubric_count = 1;

	#Идем по подрубрикам
	foreach my $rubric (@{$item->{'rubrics'}})
	{
	    println("\t#: $rubric_count", 'bold green');
	    printlnEx("\tName: " . $rubric->{'name'}, $COLOR_NAME);
	    println("\tID: " . $rubric->{'id'});
	    println("\tParent ID: " . $rubric->{'parent_id'});
	    println("\torg_count: " . $rubric->{'org_count'});
	    println("\tbranch_count: " . $rubric->{'branch_count'});
	    println("\tType: " . $rubric->{'type'});

            println('-' x $COUNT_COLS, $COLOR_DELIMITER);

	    $rubric_count++;

	    #Идем и собираем информацию о компаниях
	    get_data_about_compamy($rubric->{'id'}, $item->{'name'}, $rubric->{'name'});
	}
        println('-' x $COUNT_COLS, $COLOR_DELIMITER);
		
		#Пишем каждую рубрику в отдельный файл если EXCEL_FULL => 0
		if (!EXCEL_FULL)
		{
			#т.к символ '/' в названии файла недопустим то меняем его на '.'
			$item_name =~ tr/\//./;
            
			save_to_excel("$item_count) $item_name _" . get_date() . ".xlsx") if SAVE_TO_EXCEL;
			
			$COUNT_COMPANY = 0;
			#Очищаем хэш, т.к следущая итерация будет будет с нуля его заполнять
			undef %MAIN_DATA;
		}
		
		$item_count++;
		#$MAIN_DATA{'result'}{'total'} = $COUNT_COMPANY;
		
    }
	
	#Сохраняем в файл если задано
	my $file_json = get_name('2gis', 'json');
	save_json_to_file($file_json) if SAVE_JSON_TO_FILE;
	
	#Сохраняем все данные в один файл
	if (EXCEL_FULL)
	{
		#Сохраняем в Excel если задано
		my $file_xlsx = get_name('2gis', 'xlsx');
		save_to_excel($file_xlsx) if SAVE_TO_EXCEL;
	}

	undef %MAIN_DATA;
}

sub get_city_by_id
{
my $id = shift;
}

sub get_id_of_city
{
    my $city = shift;

    print "City: $city\n";

    my $id_city;

    my $url = "http://catalog.api.2gis.ru/geo/search?q=$CITY&version=1.3&key=$KEY";

    my $response = get($url);

    my $decoded = decode_json($response);

    my @result = @{$decoded->{'result'}};

    foreach my $id (@result)
    {
	$id_city = $id->{'project_id'}; 
    }

    print "ID City: $id_city\n";

    return $id_city;
}


sub get_data_about_compamy
{
    my $rubric_id = $_[0];
    my $general_rubric = $_[1];
    my $sub_rubric = $_[2];

    my $page = 1;
	
	#Количество выводимых компании на странице
    my $page_size = '12';

    my $item_count = 1;

    while(1)
    {
	my $url = $GENERAL_URL . "branch/list?page=$page&page_size=$page_size&rubric_id=$rubric_id&region_id=$REGION_ID&fields=items.region_id%2Citems.contact_groups%2Citems.address&key=$KEY";

	my $response = get($url);

	my $decoded = decode_json($response);

	my $response_code = $decoded->{'meta'}{'code'};

	#Идем по страницам пока получаем в ответ 200
	#Описание кодов ответа смотрим выше response_code comment
	if ($response_code == 200)
	{
	    print "Total: " . $decoded->{'result'}{'total'} . "\n";

	    binmode STDOUT, ":utf8";

	    my @items = @{$decoded->{'result'}{'items'}};

	    foreach my $item (@items)
	    {
		{#no warnings
		    no warnings 'uninitialized';

		    println("\t\t#: $item_count", 'bold green');
		    printlnEx("\t\tName: " . $item->{'name'}, $COLOR_NAME);
                    println("\t\tID: " . $item->{'id'});

                    my $api_link = $GENERAL_URL . "branch/get?id=" . $item->{'id'} . "&key=$KEY";
                    println("\t\tAPI link: $api_link");

		    println( "\t\tRegion ID: " . $item->{'region_id'});

		    println("\t\tAddress: " . $item->{'address_name'});
		    println("\t\tAddress comment: " . $item->{'address_comment'});
		    println("\t\tBuilding name: " . $item->{'address'}{'building_name'});
		    println("\t\tPostcode: " . $item->{'address'}{'postcode'});
		    println("\t\tArticle: " . delete_tags($item->{'ads'}{'article'}));
		    println("\t\tText: " . $item->{'ads'}{'text'});
		    println("\t\tGeneral rubric: $general_rubric");
		    println("\t\tSubrubric: $sub_rubric");

                    my $phone = get_phone($item->{'contact_groups'});
		    println("\t\tPhone: " . get_phone($item->{'contact_groups'}));

                    my $web_site = get_website($item->{'contact_groups'});
		    println("\t\tWeb site: " . get_website($item->{'contact_groups'}));

                    my $email = get_email($item->{'contact_groups'});
		    println("\t\tEmail: " . get_email($item->{'contact_groups'}));
                    println('-' x $COUNT_COLS, $COLOR_DELIMITER);

		    #Заполняем полученными даными
                    fill_hash($COUNT_COMPANY,
				$item->{'name'},
				$item->{'id'},
				$api_link,
				$item->{'region_id'},
				$item->{'address_name'},
				$item->{'address_comment'},
				$item->{'address'}{'building_name'},
				$item->{'address'}{'postcode'},
				$phone,
                                $web_site,
                                $email,
				delete_tags($item->{'ads'}{'article'}),
				$item->{'ads'}{'text'},
				$general_rubric,
				$sub_rubric
				);
			$COUNT_COMPANY++;
		}#end no warnings
		$item_count++;
	    }#end foreach
	    $page++;
           $MAIN_DATA{'result'}{'total'} = $COUNT_COMPANY;
	}
	else
	{
	    return;
	}
    }#end while
}

sub get_phone
{
    my @json_contacts = shift;

    my $phone = '';

    foreach my $contacts1 (@json_contacts)
    {
	foreach my $contact1 (@{$contacts1})
	{
	    my @contacts2 = @{$contact1->{'contacts'}};

            my $count = scalar @contacts2;

	    #Пробегаем по контактам
	    #text - очень удобен для для чтения, например: +7 (383) 363-59-68
	    #value - если планируется использовать автонабор, например: +73833635968
	    #comment - комментарии к номеру
	    foreach my $contact2 (@contacts2)
	    {
		if ($contact2->{'type'} eq 'phone')
		{
		    $phone .= $contact2->{'text'} . "\n" if not $contact2->{'comment'};
		    $phone .= $contact2->{'text'} . ' ' . $contact2->{'comment'} . "\n" if $contact2->{'comment'};
		}
	    }
	}
    }
    return $phone;
}

sub get_website
{
    my @json_contacts = shift;

    my $website = '';

    foreach my $contacts1 (@json_contacts)
    {
        foreach my $contact1 (@{$contacts1})
        {
            my @contacts2 = @{$contact1->{'contacts'}};

            foreach my $contact2 (@contacts2)
            {
                if ($contact2->{'type'} eq 'website')
                {
		    $website .= $contact2->{'text'};
		    $COUNT_COMPANY_WEBSITE++;
                }
            }
        }
    }
    return $website;
}

sub get_email
{
    my @json_contacts = shift;

    my $email = '';

    foreach my $contacts1 (@json_contacts)
    {
        foreach my $contact1 (@{$contacts1})
        {
            my @contacts2 = @{$contact1->{'contacts'}};

            foreach my $contact2 (@contacts2)
            {
                if ($contact2->{'type'} eq 'email')
                {
		    $email .= $contact2->{'text'};
		    $COUNT_COMPANY_EMAIL++;
                }
            }
        }
    }
    return $email;
}

sub send_report_about_parsing
{

my $msg = MIME::Lite->new(
	From     => FROM,
	To       => RECIPIENT,
	Subject  => SUBJECT,
	Type     => 'multipart/related'
);

my $body = qq{ <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
					<html xmlns="http://www.w3.org/1999/xhtml">
					<head>
					<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
						<style type="text/css">
						p {
							line-height: 1.0;
						}
						</style>
					</head>
					<body>
					</body>
					    </html>
					};

$msg->attach(
	Type => 'text/html',
	Data => $body
);

$msg->send('smtp', SMTP_SERVER, Debug => EMAIL_DEBUG);
}

sub println
{
    my ($text, $color) = @_;

    print color('reset');

    if (defined $color)
    {
        print color($color);
        print("$text\n");
        return;
    }

    print("$text\n");
    print color('reset');
}

sub printlnEx
{
    my ($text, $color) = @_;

    print colored($text, $color), "\n";
}

sub delete_tags
{
    my $text = shift;

    if (defined $text)
    {
	my $hs = HTML::Strip->new();

	my $plain_text = $hs->parse($text);

	$hs->eof;

	$plain_text =~ s!<br />&#149;&nbsp;!!sg;

	return $plain_text;
    }
    return '';
}

sub fill_hash
{
    my ($count,
	$name,
	$id,
	$api_link,
	$region_id,
	$address,
	$address_comment,
	$building_name,
	$postcode,
	$phone,
	$web_site,
	$email,
	$article,
	$text,
	$general_rubric,
	$sub_rubric) = @_;

    $MAIN_DATA{'result'}{'items'}[$count]{'name'} = $name;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'id'} = $id;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'api_link'} = $api_link;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'region_id'} = $region_id;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'address'} = $address;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'address_comment'} = $address_comment;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'building_name'} = $building_name;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'postcode'} = $postcode;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'phone'} = $phone;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'web_site'} = $web_site;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'email'} = $email;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'article'} = $article;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'text'} = $text;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'general_rubric'} = $general_rubric;
    $MAIN_DATA{'result'}{'items'}[$count]{'attributes'}{'sub_rubric'} = $sub_rubric;

}

sub generate_string
{
    my $length_of_randomstring = $_[0];

    my @chars = ('a'..'z','A'..'Z','0'..'9','_');

    my $random_string;

    foreach (1..$length_of_randomstring) 
    {
	$random_string .= $chars[rand @chars];
    }

    return $random_string;
}

sub fill_hash_test
{
    my $count = shift;

    my $progress = Term::ProgressBar->new ({count => $count ,name => "Заполнение тестовых данных ($count строк)"});

    for (my $i = 0; $i <= $count; $i++)
    {

	$MAIN_DATA{'result'}{'items'}[$i]{'name'} = generate_string(20);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'id'} = generate_string(30);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'api_link'} = generate_string(40);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'region_id'} = generate_string(1);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'address'} = generate_string(15);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'address_comment'} = generate_string(10);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'building_name'} = generate_string(6);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'postcode'} = generate_string(5);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'phone'} = generate_string(12);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'web_site'} = generate_string(10);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'email'} = generate_string(16);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'article'} = generate_string(150);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'text'} = generate_string(60);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'general_rubric'} = generate_string(60);
	$MAIN_DATA{'result'}{'items'}[$i]{'attributes'}{'sub_rubric'} = generate_string(60);

	$progress->update($_);
    }
    $MAIN_DATA{'result'}{'total'} = $count;
    print 'Total size hash (before): ' . convert_to_mb(total_size(\%MAIN_DATA)). " MB\n";

    #%MAIN_DATA = encode_json(%MAIN_DATA);
    my $encoded = encode_json(\%MAIN_DATA);
    print 'Total size hash (after): ' . convert_to_mb(total_size($encoded)). " MB\n";
}

sub save_json_to_file
{
    my $file = shift;

    open my $fh, ">", $file;

    print $fh encode_json(\%MAIN_DATA);

    close $fh;

    print "Файл \'$file\': СОХРАНЕН\n";

    archive_file($file) if ARCHIVE_JSON;
}

sub save_to_excel
{
    my $file = shift;

    #my $date = strftime '%Y%m%d', localtime;

	my $path = PATH_FOR_SAVING . 'city_' . $REGION_ID . '_date_' . get_date() . '/';
	
	#Создаем папку, например:  data_20220115
	mkdir $path;
	
	#my $file = get_name('2gis', 'xlsx');
    #my $file = '2gis_city-' . $REGION_ID . '_date-' . "$date.xlsx";

    my $workbook  = Excel::Writer::XLSX->new($path . $file);
    my $worksheet = $workbook->add_worksheet($REGION_ID);

    $workbook->set_properties(
				title    => 'Data',
				author   => 'Noname',
				comments => 'Created by Perl and Excel::Writer::XLSX',
    );

    my $format_header = $workbook->add_format(border => 2);

    #Шрифт для заголовка
    $format_header->set_bold();
    $format_header->set_color('red');
    $format_header->set_size(18);
    $format_header->set_font('Cambria');

    #Выравнивание заголовка
    $format_header->set_align('center');

    #Заливка заголовка
    $format_header->set_bg_color('#FFFFCC');

    $worksheet->write("A1", decode('UTF-8', 'Название'), $format_header );
    $worksheet->write("B1", decode('UTF-8', 'ID'), $format_header );
    $worksheet->write("C1", decode('UTF-8', 'Адрес'), $format_header);
    $worksheet->write("D1", decode('UTF-8', 'Комментарии'), $format_header);
    $worksheet->write("E1", decode('UTF-8', 'Название здания'), $format_header);
    $worksheet->write("F1", decode('UTF-8', 'Телефон'), $format_header );
    $worksheet->write("G1", decode('UTF-8', 'Сайт'), $format_header);
    $worksheet->write("H1", decode('UTF-8', 'Email'), $format_header);
    $worksheet->write("I1", decode('UTF-8', 'Описание'), $format_header);
    $worksheet->write("J1", decode('UTF-8', 'Текст'), $format_header);
    $worksheet->write("K1", decode('UTF-8', 'Индекс'), $format_header);
    $worksheet->write("L1", decode('UTF-8', 'Рубрика'), $format_header);
    $worksheet->write("M1", decode('UTF-8', 'Подрубрика'), $format_header);

    #Пишем комментарии
    $worksheet->write_comment( 'A1', decode('UTF-8', 'Название компании'));
    $worksheet->write_comment( 'B1', decode('UTF-8', 'Уникальный ID города'));
    $worksheet->write_comment( 'C1', decode('UTF-8', 'Адрес'));
    $worksheet->write_comment( 'D1', decode('UTF-8', 'Комментарии к адресу'));
    $worksheet->write_comment( 'E1', decode('UTF-8', 'Название здания где расположена компания, например: ТЦ \'Спутник\''));
    $worksheet->write_comment( 'F1', decode('UTF-8', 'Телефон(ны) компании'));
    $worksheet->write_comment( 'G1', decode('UTF-8', 'Веб сайт компании'));
    $worksheet->write_comment( 'H1', decode('UTF-8', 'Электронный адрес компании'));
    $worksheet->write_comment( 'I1', decode('UTF-8', 'This is a comment for Article'));
    $worksheet->write_comment( 'J1', decode('UTF-8', 'This is a comment for Text'));
    $worksheet->write_comment( 'K1', decode('UTF-8', 'Почтовый индекс'));
    $worksheet->write_comment( 'L1', decode('UTF-8', 'Рубрика, например: \'Автосервис / Автотовары\''));
    $worksheet->write_comment( 'M1', decode('UTF-8', 'Подрубрика, например: \'Автозапчасти для иномарок\''));


    #Закрепить первую строку
    $worksheet->freeze_panes(1, 0);

    my $format = $workbook->add_format(border => 1);

    #Шрифт для данных
    $format->set_color('black');
    $format->set_size(14);
    $format->set_font('Cambria');
    $format->set_text_wrap();

    #Выравнивание
    $format->set_align('left');
    $format->set_align('vcenter');

    #Устанавливаем ширину ячеек
    $worksheet->set_column('A:A', 25);
    $worksheet->set_column('B:B', 13);
    $worksheet->set_column('C:C', 40);
    $worksheet->set_column('D:D', 35);
    $worksheet->set_column('E:E', 35);
    $worksheet->set_column('F:F', 35);
    $worksheet->set_column('G:G', 20);
    $worksheet->set_column('H:H', 15);
    $worksheet->set_column('I:I', 25);
    $worksheet->set_column('J:J', 15);
    $worksheet->set_column('K:K', 20);
    $worksheet->set_column('L:L', 40);
    $worksheet->set_column('M:M', 40);

    #Включаем автофильтр
    $worksheet->autofilter('A1:M1');

    my $total = $MAIN_DATA{'result'}{'total'};

    my $progress = Term::ProgressBar->new ({count => $total, name => "Запись данных ($total строк)"});

    foreach my $row (0..$total)
    {

	$worksheet->write($row+1, 0, $MAIN_DATA{'result'}{'items'}[$row]{'name'}, $format);
	$worksheet->write($row+1, 1, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'region_id'}, $format);
	$worksheet->write($row+1, 2, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'address'}, $format);
	$worksheet->write($row+1, 3, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'address_comment'}, $format);
	$worksheet->write($row+1, 4, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'building_name'}, $format);
	$worksheet->write($row+1, 5, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'phone'}, $format);
	$worksheet->write($row+1, 6, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'web_site'}, $format);
	$worksheet->write($row+1, 7, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'email'}, $format);
	$worksheet->write($row+1, 8, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'article'}, $format);
	$worksheet->write($row+1, 9, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'text'}, $format);
	$worksheet->write($row+1, 10, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'postcode'}, $format);
	$worksheet->write($row+1, 11, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'general_rubric'}, $format);
	$worksheet->write($row+1, 12, $MAIN_DATA{'result'}{'items'}[$row]{'attributes'}{'sub_rubric'}, $format);

	$progress->update($_);

	#После того как записали в Excel - удаляем, т.к планируется работа на ведре
	delete $MAIN_DATA{'result'}{'items'}[$row];

	#for debug
	#print "Total size hash: " . total_size(\%MAIN_DATA) . ', in MB: ' . convert_to_mb(total_size(\%MAIN_DATA)). "\n";
        #print "Size hash:" . size(\%MAIN_DATA) . "\n";
    }

    $workbook->close;

    print "Файл \'$file\': СОХРАНЕН\n";

    undef $workbook;

    archive_file($file) if ARCHIVE_EXCEL;
}

sub convert_to_mb
{
    my $size = shift;

    return ceil($size / (1024 * 1024));
}

sub archive_file
{
    my $file = shift;

    my $zip = Archive::Zip->new();

    my $file_member = $zip->addFile($file);

    #my $archive_name = &delete_extension($file) . '.zip';
    my $archive_name = $file . '.zip';

    unless ($zip->writeToFileNamed($archive_name) == AZ_OK)
    {
	die 'Ошибка архивирования';
    }

    delete_file($file);

    print "Архивация \'$archive_name\': УСПЕШНО\n";
}

sub delete_extension
{
    my $name = shift;

    $name =~ s{\.[^.]+$}{};

    return $name;
}

sub delete_file
{
    my $file = shift;

    unlink $file;

    print "Удаление файла \'$file\': УСПЕШНО\n";
}

sub print_summary_info
{
    printlnEx('******************SUMMARY INFORMATION**********************', 'red on_bright_yellow');

    println("General rubric: $COUNT_GENERAL_RUBRIC");
    println("Rubric: $COUNT_RUBRIC");
    println("Company: $COUNT_COMPANY");

    my $percent_company_email = int($COUNT_COMPANY_EMAIL*100/$COUNT_COMPANY);
    println("Company with email: $COUNT_COMPANY_EMAIL, ($percent_company_email%)");

    my $percent_company_website = int($COUNT_COMPANY_WEBSITE*100/$COUNT_COMPANY);
    println("Company with website: $COUNT_COMPANY_WEBSITE, ($percent_company_website%)");
}

sub main
{
	system('clear');
	
	parse_argv();
			
	#Если тест то заполняем тестовыми данными и завершаем работу скрипта
	do_test() if TEST;

	if (!defined $KEY)
	{
		printlnEx('!!! Нужен ключ !!!', 'red');
		
		exit 1;
	}
	
	my $time_start = strftime "%Y-%m-%d %H:%M:%S", localtime;
	
    $COUNT_COLS = `tput cols`;

	#Собственно сбор данных
    get_data();

	#Вывод информации при отладке
    print_summary_info() if DEBUG;
	
	#Посылает отчет о парсинге
	#send_report_about_parsing() if REPORT_VIA_EMAIL;

    my $time_end = strftime "%Y-%m-%d %H:%M:%S", localtime;

    println("Начало выполнения скрипта $0: " . "$time_start") if DEBUG;
    println("Завершение выполнения скрипта $0: " . "$time_end") if DEBUG;
    printlnEx('******************END**********************', 'red on_bright_yellow');
}
