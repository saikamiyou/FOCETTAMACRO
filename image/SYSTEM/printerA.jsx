// argument.info �t�@�C���̃p�X���w��
// CSV�t�@�C���̃p�X���w��

var argumentFilePath = new File(File($.fileName).parent + "/argument.info");

// argument.info �t�@�C�����J��
var argumentFile = new File(argumentFilePath);
if (argumentFile.exists) {

    argumentFile.open('r');
    
    // �e�s�������ǂݍ���
    while (!argumentFile.eof) {
        var filePath = argumentFile.readln();
        
		// trim()�̑�ւƂ��āA�蓮�őO��̋󔒂��폜
        filePath = filePath.replace(/^\s+|\s+$/g, '');

        if (filePath) {
            var file = new File(filePath);
            if (file.exists) {
                // �h�L�������g���J��
                var doc = open(file);

                // ����ݒ�
                function printDocument() {
                    try {
                        // Photoshop�ł̈����Action�i�A�N�V�����j���g���čs���܂�
                        var idPrnt = charIDToTypeID( "Prnt" );
                        var desc = new ActionDescriptor();
                        desc.putEnumerated( charIDToTypeID( "Nm  " ), charIDToTypeID( "PntC" ), charIDToTypeID( "Cnvs" ) );
                        executeAction( idPrnt, desc, DialogModes.NO );
                        
                        alert("����W���u�����M����܂����B");
                    } catch (e) {
                        alert("������ɃG���[���������܂���: " + e.message);
                    }
                }

                // ��������s
                printDocument();

                // �h�L�������g�����i�ۑ����Ȃ��j
                doc.close(SaveOptions.DONOTSAVECHANGES);
            } else {
                alert("�w�肳�ꂽ�t�@�C����������܂���: " + filePath);
            }
        }
    }
    
    // �t�@�C�������
    argumentFile.close();
} else {
    alert("argument.info �t�@�C����������܂���: " + argumentFilePath);
}