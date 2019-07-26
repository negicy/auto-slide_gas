  /* このApplicationは, GoogleDriveのフォルダ内の画像から自動でスライドを作ります。
     また, 選んだフォルダの名前をスライドのタイトルとして1枚目のスライドに表示させることができます。
     対象ユーザー: 活動報告など, 写真メインのスライドショーを自動で作成したい方。 */

  function myFunction() {
  //変数prsにこのslideを代入
  var prs = SlidesApp.getActivePresentation();
  //写真の入ったフォルダを取得 by ID
  //変数folは取得したファイルを代入する変数
  //fol_id:写真の入ったフォルダIDを代入すry変数IDを代入する変数
  var fol_id = "";
  var fol = DriveApp.getFolderById(fol_id); //※フォルダID要変更
  //フォルダの中の写真ファイルを一括取得　※fol!=DriveApp
  var files = fol.getFiles();
  //変数filesについて繰り返し
  //hasNext methodで取得したファイルたちfilesの中にまだ処理していないfileがあるか確認,
  // file.hasNext==trueの間はwhileで繰り返し
  while(files.hasNext()) {
    var file = files.next();
    //JPEG,GIF,PNG画像のみ処理
    //正規表現でfileを判別
    //getMimeType:ファイル識別子を指定して取得, matchで具体的に正規表現で判別
    if(file.getMimeType().match(/^image\/(?:JPG|jpeg|png|jpg)$/i)) {
      try {
        addImageSlide(prs, file.getBlob());
        NameFunction();
      } catch(e) {
        Logger.log(file.getName() + ":" + e.message);
      }
    }
  }
  SlidesApp.getUi().alert("処理が終了しました。");
}
 
//空白スライドを追加し画像を挿入
function addImageSlide(prs, imageBlob) {
  //空白のslideを挿入
  var slide = prs.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  
  var image = slide.insertImage(imageBlob);
  
  var imgWidth = image.getWidth();
  var imgHeight = image.getHeight();
  
  var pageWidth = prs.getPageWidth();
  var pageHeight = prs.getPageHeight();
  
  var newX = (pageWidth / 2) - (imgWidth / 2);
  var newY = (pageHeight / 2) - (imgHeight / 2);
  image.setLeft(newX).setTop(newY); //画像中央揃え

}
function NameFunction() {
  //title: フォルダ名を1枚目のスライドに表示したいフォルダIDを代入する変数
  var title = DriveApp.getFolderById("").getName();
  var slide = SlidesApp.getActivePresentation().getSlides()[0]; //スライド1枚目
  //テキストボックス追加(insertShape(shapeType, left, top, width, height))
  var shape = slide.insertShape(SlidesApp.ShapeType.TEXT_BOX, 100, 100, 600, 120);
  var txtRng = shape.getText();
  if(title=="title"){
    Logger.log("終了しました")
    return;
  }else{
  txtRng.setText(title); //テキストボックスの文字設定
  //文字装飾：
  txtRng.getTextStyle().setBold(true)
                       .setFontFamily("Indie Flower")
                       .setFontSize(128);
  //中央揃え
  txtRng.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }
}



