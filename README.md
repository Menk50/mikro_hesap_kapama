# mikro_hesap_kapama
Mikro ERP entegreli Pazar Yeri Bazlı Cari Hesap Kapama Uygulamasının bazı örnek kodlarını paylaşıyorum. tüm proje kodları Private rep olarak beklemektedir.
Full Mikro Entegreli bir şekilde çalışmaktadır.

        SqlCommand count_sorgu = new SqlCommand(@"SELECT Count(cha_evrakno_seri) AS 'count' from CARI_HESAP_HAREKETLERI where cha_evrakno_seri ='"+evrak_seri_N11+"'",baglanti);
        SqlCommand say = new SqlCommand(@"SELECT max(s.cha_evrakno_sira) from CARI_HESAP_HAREKETLERI s where s.cha_evrakno_seri = '" + evrak_seri_N11 + "'", baglanti);
        baglanti_kont();
        count_sorgu.ExecuteNonQuery();
        SqlDataReader iki = count_sorgu.ExecuteReader();
        if (iki.Read())
        {
            say_count = iki["count"].ToString();
            
        }
        //label18.Text = " say: "+say_count.ToString();
        iki.Close();
        if (Convert.ToInt32(say_count)==0)
        {
            label18.Text = evrak_seri_N11 +" Seri Numarası ile ilk Evrak Girişidir.";
            label31.Text = evrak_seri_N11 + " Seri Numarası ile ilk Evrak Girişidir.";

            count = +1;
        }
        else if (Convert.ToInt32(say_count)!=0)
        {
            SqlDataReader oku = say.ExecuteReader();
            while (oku.Read())
            {
                // countt = oku.GetString(0);
                count = oku.GetInt32(0);
            }
        }
        // baglanti.Open();
       
        // label2.Text = oku.ToString();
        //  count = Convert.ToInt32(oku);
        //label2.Text = label2.Text + " " + count;
        if (count >= evrak_sira)
        {
            evrak_sira = count + 1;
        }

        baglanti_kont();
for (int i = 0; i < kap_evrak_say; i++) {
                baglanti_kont();
                toplam += Convert.ToDouble(kapatilacak_tablo.Rows[i].Cells[3].Value);

                insertkomut = new SqlCommand("INSERT INTO[dbo].[CARI_HESAP_HAREKETLERI]([cha_DBCno],[cha_SpecRecNo],[cha_iptal],[cha_fileid],[cha_hidden],[cha_kilitli],[cha_degisti],[cha_CheckSum],[cha_create_user],[cha_create_date],[cha_lastup_user],[cha_lastup_date],[cha_special1],[cha_special2],[cha_special3],[cha_firmano],[cha_subeno],[cha_evrak_tip],[cha_evrakno_seri],[cha_evrakno_sira],[cha_satir_no],[cha_tarihi],[cha_tip],[cha_cinsi],[cha_normal_Iade],[cha_tpoz],[cha_ticaret_turu],[cha_belge_no],[cha_belge_tarih],[cha_aciklama],[cha_satici_kodu],[cha_EXIMkodu],[cha_projekodu],[cha_yat_tes_kodu],[cha_cari_cins],[cha_kod],[cha_ciro_cari_kodu],[cha_d_cins],[cha_d_kur],[cha_altd_kur],[cha_grupno],[cha_srmrkkodu],[cha_kasa_hizmet],[cha_kasa_hizkod],[cha_karsidcinsi],[cha_karsid_kur],[cha_karsidgrupno],[cha_karsisrmrkkodu],[cha_miktari],[cha_meblag],[cha_aratoplam],[cha_vade],[cha_Vade_Farki_Yuz],[cha_ft_iskonto1],[cha_ft_iskonto2],[cha_ft_iskonto3],[cha_ft_iskonto4],[cha_ft_iskonto5],[cha_ft_iskonto6],[cha_ft_masraf1],[cha_ft_masraf2],[cha_ft_masraf3],[cha_ft_masraf4],[cha_isk_mas1],[cha_isk_mas2],[cha_isk_mas3],[cha_isk_mas4],[cha_isk_mas5],[cha_isk_mas6],[cha_isk_mas7],[cha_isk_mas8],[cha_isk_mas9],[cha_isk_mas10],[cha_sat_iskmas1],[cha_sat_iskmas2],[cha_sat_iskmas3],[cha_sat_iskmas4],[cha_sat_iskmas5],[cha_sat_iskmas6],[cha_sat_iskmas7],[cha_sat_iskmas8],[cha_sat_iskmas9],[cha_sat_iskmas10],[cha_yuvarlama],[cha_StFonPntr],[cha_stopaj],[cha_savsandesfonu],[cha_avansmak_damgapul],[cha_vergipntr],[cha_vergi1],[cha_vergi2],[cha_vergi3],[cha_vergi4],[cha_vergi5],[cha_vergi6],[cha_vergi7],[cha_vergi8],[cha_vergi9],[cha_vergi10],[cha_vergisiz_fl],[cha_otvtutari],[cha_otvvergisiz_fl],[cha_oiv_pntr],[cha_oivtutari],[cha_oiv_vergi],[cha_oivergisiz_fl],[cha_fis_tarih],[cha_fis_sirano],[cha_trefno],[cha_sntck_poz],[cha_reftarihi],[cha_istisnakodu],[cha_pos_hareketi],[cha_meblag_ana_doviz_icin_gecersiz_fl],[cha_meblag_alt_doviz_icin_gecersiz_fl],[cha_meblag_orj_doviz_icin_gecersiz_fl],[cha_sip_uid],[cha_kirahar_uid],[cha_vardiya_tarihi],[cha_vardiya_no],[cha_vardiya_evrak_ti],[cha_ebelge_turu],[cha_tevkifat_toplam],[cha_ilave_edilecek_kdv1],[cha_ilave_edilecek_kdv2],[cha_ilave_edilecek_kdv3],[cha_ilave_edilecek_kdv4],[cha_ilave_edilecek_kdv5],[cha_ilave_edilecek_kdv6],[cha_ilave_edilecek_kdv7],[cha_ilave_edilecek_kdv8],[cha_ilave_edilecek_kdv9],[cha_ilave_edilecek_kdv10],[cha_e_islem_turu],[cha_fatura_belge_turu],[cha_diger_belge_adi],[cha_uuid],[cha_adres_no],[cha_vergifon_toplam],[cha_ilk_belge_tarihi],[cha_ilk_belge_doviz_kuru],[cha_HareketGrupKodu1],[cha_HareketGrupKodu2],[cha_HareketGrupKodu3],[cha_ebelgeno_seri],[cha_ebelgeno_sira],[cha_hubid],[cha_hubglbid],[cha_disyazilimid],[cha_eticaret_kanal_kodu],[cha_disyazilim_tip],[cha_bsba_e_belge_mi])     VALUES('0', '0', '0', '51', '0', '0', '0', '0', '" + Properties.Settings.Default.Mikro_olusturan_kullanici + "', '" + tarihh + "', '" + Properties.Settings.Default.Mikro_son_kullanici + "', '" + tarihh + "', '', '', '', '0', '0', '33', '" + evrak_seri_N11 + "', '" + evrak_sira + "', '" + i + "', '" + tarihh + "', '1', '5', '0', '0', '0', '', '" + tarihh + "', '" + evrak_seri_N11 + "', '', '', '', '', '0', '" + kapatilacak_tablo.Rows[i].Cells[0].Value + "', '', '0', '1', '11.02', '0', '1', '0', '', '0', '1', '0', '', '0', '" + kapatilacak_tablo.Rows[i].Cells[3].Value + "', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '2021-11-19 00:00:00.000', '7', '', '0', '1899-12-30 00:00:00.000', '0', '0', '0', '0', '0', '00000000-0000-0000-0000-000000000000', '00000000-0000-0000-0000-000000000000', '1899-12-30 00:00:00.000', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '0', '', '', '0', '0', '1899-12-30 00:00:00.000', '0', '', '', '', '', '0', '', '', '', '0', '0', '0')", this.baglanti);
                int k = insertkomut.ExecuteNonQuery();
                baglanti_kont();

            }
<baglanti_kont> private void baglanti_kont() {

        if (baglanti.State != ConnectionState.Open)
        {

            baglanti.Open();
        }
        else if (baglanti.State != ConnectionState.Closed)
        {

            baglanti.Close();
        }
    }
</baglanti_kont>
