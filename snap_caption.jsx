/*
snap_caption.jsx
(c)2007-2009 www.seuzo.jp

キャプションを画像にスナップさせます。
簡単な使い方説明：http://jp.youtube.com/watch?v=a8ar4Ydq0ww

・開発＆動作環境
Mac Pro Quad 3GHz/4GB Memory
-MacOS X10.5.6
-InDesign CS4_J（6.0.1）

・History
2007-12-16	ver.0.1	とりあえず横組キャプションしか考えてません。
2007-12-17	ver.0.2	「設定」にmy_text_align_topを追加
2007-12-18	ver.0.3	画像がグループ化されていても動作するようにした。グラフィックフレームが楕円または多角形でも動作するようにした。「設定」にmy_text_fitを追加。テキストフレームの上辺がグラフィックフレームの底辺のわずかに（my_distance値の２倍分）上にあっても、下揃えのキャプションと判定することにした。
2007-12-19	ver.0.4	my_text_fitをテキストフレームがオーバーセットしている時のみ動作するようにした。
2007-12-21	ver.0.5	上付きキャプションに対応した。全体的にコードの整理、、、整理したはずなのに、だんだんジャミってkt..
2007-12-23	ver.0.6	クリスマス前で暇なので、コードの生理。血だらけ。
2007-12-24	ver.0.7	いぜんとして暇なので（クライアントさまはクリスマスイブをエンジョイしてるようだ）、縦組キャプションにも対応した。クリスマスを撃ち殺せバージョン（そしてやはり血だらけ）。
2008-03-16	ver.0.8	ページルーラーの開始位置が「スプレッド」以外になっていたら、「スプレッド」に一時的に変更するようにした。オブジェクトがロックされている時、エラーにするようにした。グループ中のひとつが選択されている時、エラーになるようにした。
2009-04-18	ver.0.9	InDesign CS4対応。スプレッドが回転表示しているときは、スクリプトの実行を中止するようにした。
*/


////////////////////////////////////////////設定
const my_distance = 1; //画像とキャプションの間隔（単位は環境設定に依存）
const my_width_adjustment = true; //キャプションが画像の下または上にあるとき、テキストフレームの幅を画像にあわせるかどうか
const my_text_align_bottom = true; //キャプションが画像の右（左）にあり、かつ画像の下側に近いとき、または上付きのキャプションのとき、テキストフレームのそろえを下にするかどうか
const my_text_align_top = true; //キャプションを画像の下、または画像の上辺にフィットさせるとき、テキストフレームのそろえを強制的に上揃えにするかどうか
const my_text_fit = true; //テキストフレームがオーバーセットしているとき、テキストフレームに対して「フレームを内容にあわせる」を実行するかどうか


////////////////////////////////////////////エラー処理 
function myerror(mess) { 
  if (arguments.length > 0) { alert(mess); }
  exit();
}


////////////////////////////////////////////スプレッドの回転角度を調べる 
function spread_angle(spread_obj) {
	var my_document = app.activeDocument;
	var my_old_ruler_origin = false;//
	if (my_document.viewPreferences.rulerOrigin != 1380143215) {//not page
		my_old_ruler_origin = my_document.viewPreferences.rulerOrigin;//current setting
		my_document.viewPreferences.rulerOrigin = 1380143215;//change
	}
	var my_old_zeroPoint = false;
	if (my_document.zeroPoint != [0, 0]) {
		my_old_zeroPoint = my_document.zeroPoint;
		my_document.zeroPoint = [0,0];
	}

	var my_page = spread_obj.pages[0];
	var my_page_bounds = my_page.bounds;
	var my_angle = -1;
	if((my_page_bounds[0] == 0) && (my_page_bounds[1] == 0)) {
		my_angle = 0;
	} else if ((my_page_bounds[0] == 0) && (my_page_bounds[3] == 0)) {
		my_angle = 90;
	} else if ((my_page_bounds[2] == 0) && (my_page_bounds[3] == 0)) {
		my_angle = 180;
	} else if ((my_page_bounds[1] == 0) && (my_page_bounds[2] == 0)) {
		my_angle = 270;
	}

	if(my_old_ruler_origin) {my_document.viewPreferences.rulerOrigin = my_old_ruler_origin}
	if(my_old_zeroPoint) {my_document.zeroPoint = my_old_zeroPoint}
	return my_angle;
}

////////////////////////////////////////////オブジェクトの座標と幅、高さを得る 
function get_bounds(my_obj) {
	var my_obj;
	var tmp_hash = new Array();
	var my_obj_bounds = my_obj.visibleBounds; //オブジェクトの大きさ（線幅を含む）
	tmp_hash["y1"] = my_obj_bounds[0];
	tmp_hash["x1"] = my_obj_bounds[1];
	tmp_hash["y2"] = my_obj_bounds[2];
	tmp_hash["x2"] = my_obj_bounds[3];
	tmp_hash["w"] = tmp_hash["x2"] - tmp_hash["x1"]; //幅
	tmp_hash["h"] = tmp_hash["y2"] - tmp_hash["y1"]; //高さ
	return tmp_hash; //ハッシュで値を返す
}

////////////////////////////////////////////テキストフレームに「フレームを内容にあわせる」を実行し、下／左揃えの時、位置調整をする
function expansion_frame(my_obj) {
	var my_obj;
	var my_orientation = my_obj.parentStory.storyPreferences.storyOrientation; //1986359924== 縦組、1752134266==横組
	var my_vj = my_obj.textFramePreferences.verticalJustification; //1651471469==BOTTOM_ALIGN
	var my_old_bounds = get_bounds(my_obj);
	my_obj.fit (FitOptions.frameToContent); //「フレームを内容にあわせる」を強制的に実行する
	var my_bounds = get_bounds(my_obj);
	if (my_vj == 1651471469) { //下／左揃え
		if (my_orientation == 1752134266) { //横組
			my_obj.visibleBounds = [my_old_bounds["y2"] - my_bounds["h"], my_bounds["x1"], my_old_bounds["y2"], my_bounds["x2"]];
		} else { //縦組
			my_obj.visibleBounds = [my_bounds["y1"], my_old_bounds["x1"], my_bounds["y2"], my_old_bounds["x1"] + my_bounds["w"]];
		}
	} 		
	return get_bounds(my_obj); //新しい大きさを返す
}
			

////////////////////////////////////////////以下実行ルーチン
if (app.documents.length == 0) {myerror("ドキュメントが開かれていません")}
var my_document = app.activeDocument;
var my_selection = my_document.selection;
if (my_selection.length != 2) {myerror("2つのオブジェクトを選択してください")}

//スプレッドが回転していたら、エラーで中止
var my_spread = app.layoutWindows[0].activeSpread;
if (spread_angle(my_spread) != 0) {myerror("スプレッドが回転しています。元に戻してから実行してください。")};

//オブジェクト種類の確定
if ("Rectangle, Group, Oval, Polygon".match(my_selection[0].reflect.name)) {
	var my_image_obj = my_selection[0];
	if (my_selection[1].reflect.name == "TextFrame") {
		var my_text_obj = my_selection[1];
	} else {
		myerror("グラフィックフレームとテキストフレームを選択してください");
	}
} else if (my_selection[0].reflect.name == "TextFrame") {
	var my_text_obj = my_selection[0];
	if ("Rectangle, Group, Oval, Polygon".match(my_selection[1].reflect.name)) {
		var my_image_obj = my_selection[1];
	} else {
		myerror("グラフィックフレームとテキストフレームを選択してください");
	}
} else {
	myerror("グラフィックフレームとテキストフレームを選択してください");
}

//ページルーラーの開始位置が「スプレッド」以外になっていたら、「スプレッド」に一時的に変更
var my_ruler_origin = false;//初期値はfalse
if (my_document.viewPreferences.rulerOrigin != 1380143983) {
	my_ruler_origin = my_document.viewPreferences.rulerOrigin;//現在の設定を保存
	my_document.viewPreferences.rulerOrigin = 1380143983;//「スプレッド」に一時的に変更
}

//オブジェクトの大きさ（線幅を含む）
var my_gb = get_bounds(my_image_obj);
var my_tb = get_bounds(my_text_obj);


//キャプションの場所の移動
if (my_text_obj.parentStory.storyPreferences.storyOrientation == 1752134266) {
	//【横組キャプション】
	if (my_tb["y1"] >= my_gb["y2"] - my_distance * 2) { //●キャプションが画像の下にある
		if (my_width_adjustment) { //キャプションの幅を画像にあわせる
			my_text_obj.visibleBounds = [my_gb["y2"] + my_distance, my_gb["x1"], my_gb["y2"] + my_distance + my_tb["h"], my_gb["x2"]];
		} else {
			my_text_obj.visibleBounds = [my_gb["y2"] + my_distance, my_gb["x1"], my_gb["y2"] + my_distance + my_tb["h"], my_gb["x1"] + my_tb["w"]];
		}
		if (my_text_align_top) { //テキストフレーム内を上揃え設定
			my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.topAlign;//TOP_ALIGN;
		}
	} else if (my_gb["y1"] >= my_tb["y2"]) { //●キャプションが画像の上にある
		if (my_width_adjustment) { //キャプションの幅を画像にあわせる
			my_text_obj.visibleBounds = [my_gb["y1"] - my_distance - my_tb["h"], my_gb["x1"], my_gb["y1"] - my_distance, my_gb["x2"]];
		} else {
			my_text_obj.visibleBounds = [my_gb["y1"] - my_distance - my_tb["h"], my_gb["x1"], my_gb["y1"] - my_distance, my_gb["x1"] + my_tb["w"]];
		}
		if (my_text_align_bottom) { //テキストフレーム内を下揃え設定
			my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;//BOTTOM_ALIGN;
		}
	} else if (my_tb["x2"] >= my_gb["x2"]) { //●キャプションが画像の右側にある（テキストフレームのx2がグラフィックフレームのx2よりも右にある）
		if (my_gb["y1"] + my_gb["h"] / 2 >= my_tb["y1"] + my_tb["h"] / 2) { //テキストフレームの中心がグラフィックフレームの中心より上にある
			my_text_obj.visibleBounds = [my_gb["y1"], my_gb["x2"] + my_distance, my_gb["y1"] + my_tb["h"], my_gb["x2"] + my_distance + my_tb["w"]];
			if (my_text_align_top) { //テキストフレーム内を上揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.topAlign;//TOP_ALIGN;
			}
		} else { //下にある
			my_text_obj.visibleBounds = [my_gb["y2"] - my_tb["h"], my_gb["x2"] + my_distance, my_gb["y2"], my_gb["x2"] + my_distance + my_tb["w"]];
			if (my_text_align_bottom) { //テキストフレーム内を下揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;//BOTTOM_ALIGN;
			}
		}
	} else if (my_gb["x1"] >= my_tb["x1"]) { //●キャプションが画像の左側にある（テキストフレームのx1がグラフィックフレームのx1よりも左にある）
		if (my_gb["y1"] + my_gb["h"] / 2 >= my_tb["y1"] + my_tb["h"] / 2) { //テキストフレームの中心がグラフィックフレームの中心より上にある
			my_text_obj.visibleBounds = [my_gb["y1"], my_gb["x1"] - my_distance - my_tb["w"], my_gb["y1"] + my_tb["h"], my_gb["x1"] - my_distance];
			if (my_text_align_top) { //テキストフレーム内を上揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.topAlign;//TOP_ALIGN;
			}
		} else { //下にある
			my_text_obj.visibleBounds = [my_gb["y2"] - my_tb["h"], my_gb["x1"] - my_distance - my_tb["w"], my_gb["y2"], my_gb["x1"] - my_distance];
			if (my_text_align_bottom) { //テキストフレーム内を下揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;//BOTTOM_ALIGN;
			}
		}
	} else {
		myerror("テキストフレームの位置を特定できませんでした");
	}


} else if (my_text_obj.parentStory.storyPreferences.storyOrientation == 1986359924){
	//【縦組キャプション】
	if (my_gb["x1"] >= my_tb["x2"] - my_distance * 2) { //●キャプションが画像の左側にある
		if (my_width_adjustment) { //キャプションの高さを画像にあわせる
			my_text_obj.visibleBounds = [my_gb["y1"], my_gb["x1"] -  my_distance - my_tb["w"], my_gb["y2"], my_gb["x1"] -  my_distance];
		} else {
			my_text_obj.visibleBounds = [my_gb["y1"], my_gb["x1"] -  my_distance - my_tb["w"], my_gb["y1"] + my_tb["h"], my_gb["x1"] -  my_distance];
		}
		if (my_text_align_top) { //テキストフレーム内を右揃え設定
			my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.topAlign;//TOP_ALIGN;
		}
	} else if (my_tb["x1"] >= my_gb["x2"]) { //●キャプションが画像の右側にある
		if (my_width_adjustment) { //キャプションの高さを画像にあわせる
			my_text_obj.visibleBounds = [my_gb["y1"], my_gb["x2"] +  my_distance, my_gb["y2"], my_gb["x2"] +  my_distance + my_tb["w"]];
		} else {
			my_text_obj.visibleBounds = [my_gb["y1"], my_gb["x2"] +  my_distance, my_gb["y1"] + my_tb["h"], my_gb["x2"] +  my_distance + my_tb["w"]];
		}
		if (my_text_align_bottom) { //テキストフレーム内を左揃え設定
			my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;//BOTTOM_ALIGN;
		}
	} else if (my_tb["y2"] >= my_gb["y2"]) { //●キャプションが画像の下側にある（テキストフレームのy2がグラフィックフレームのy2よりも下にある）
		if (my_tb["x1"] + my_tb["w"] / 2 >= my_gb["x1"] + my_gb["w"] / 2) { //テキストフレームの中心がグラフィックフレームの中心より右にある
			my_text_obj.visibleBounds = [my_gb["y2"] + my_distance, my_gb["x2"] - my_tb["w"], my_gb["y2"] + my_distance + my_tb["h"], my_gb["x2"]];
			if (my_text_align_top) { //テキストフレーム内を上揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.topAlign;//TOP_ALIGN;
			}
		} else { //左にある
			my_text_obj.visibleBounds = [my_gb["y2"] + my_distance, my_gb["x1"], my_gb["y2"] + my_distance + my_tb["h"], my_gb["x1"] + my_tb["w"]];
			if (my_text_align_bottom) { //テキストフレーム内を左揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;//BOTTOM_ALIGN;
			}
		}
	} else if (my_gb["y1"] >= my_tb["y1"]) { //●キャプションが画像の上側にある（テキストフレームのy1がグラフィックフレームのy1よりも上にある）
		if (my_tb["x1"] + my_tb["w"] / 2 >= my_gb["x1"] + my_gb["w"] / 2) { //テキストフレームの中心がグラフィックフレームの中心より右にある
			my_text_obj.visibleBounds = [my_gb["y1"] - my_distance - my_tb["h"], my_gb["x2"] - my_tb["w"], my_gb["y1"] - my_distance, my_gb["x2"]];
			if (my_text_align_top) { //テキストフレーム内を上揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.topAlign;//TOP_ALIGN;
			}
		} else { //左にある
			my_text_obj.visibleBounds = [my_gb["y1"] - my_distance - my_tb["h"], my_gb["x1"], my_gb["y1"] - my_distance, my_gb["x1"] + my_tb["w"]];
			if (my_text_align_bottom) { //テキストフレーム内を左揃え設定
				my_text_obj.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;//BOTTOM_ALIGN;
			}
		}
	} else {
		myerror("テキストフレームの位置を特定できませんでした");
	}
} else {
	myerror("テキストフレームの組方向を特定できませんでした");
}

//ページルーラー設定の復帰
if (my_ruler_origin) {
	my_document.viewPreferences.rulerOrigin = my_ruler_origin;
}

//オーバーセットしているテキストフレームを拡張する
if (my_text_fit && my_text_obj.overflows) {
	expansion_frame(my_text_obj);
}