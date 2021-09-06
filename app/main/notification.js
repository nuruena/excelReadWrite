const { Notification } = require( 'electron' );

// display files added notification
exports.filesAdded = ( size ) => {
    const notif = new Notification( {
        title: 'Files added',
        body: `${ size } file(s) has been successfully added.`
    } );

    notif.show();
};

// display files added notification
exports.excelNotif = ( path, name ) => {
    const notif = new Notification( {
        title: `${ name }`,
        body: `${ path }`
    } );
    notif.show();
};