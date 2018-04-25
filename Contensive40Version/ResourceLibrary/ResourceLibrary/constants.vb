﻿
Option Explicit On
Option Strict On

Namespace Contensive.Addons.ResourceLibrary
    Public Module constants
        '
        ' -- sample
        Public Const ButtonCancel As String = "Cancel"
        Public Const ButtonDelete As String = "Delete"
        Public Const ButtonApply As String = "Apply"
        Public Const ButtonClose As String = "Close"
        '
        ' -- sample
        Public Const RequestNameButton As String = "button"
        Public Const RequestNameLibraryUpload As String = "upload"
        Public Const RequestNameLibraryName As String = "name"
        Public Const RequestNameLibraryDescription As String = "description"
        Public Const xxxxx As String = "xxxx"
        '
        ' -- errors for resultErrList
        Public Enum resultErrorEnum
            errPermission = 50
            errDuplicate = 100
            errVerification = 110
            errRestriction = 120
            errInput = 200
            errAuthentication = 300
            errAdd = 400
            errSave = 500
            errDelete = 600
            errLookup = 700
            errLoad = 710
            errContent = 800
            errMiscellaneous = 900
        End Enum
        '
        ' -- http errors
        Public Enum httpErrorEnum
            badRequest = 400
            unauthorized = 401
            paymentRequired = 402
            forbidden = 403
            notFound = 404
            methodNotAllowed = 405
            notAcceptable = 406
            proxyAuthenticationRequired = 407
            requestTimeout = 408
            conflict = 409
            gone = 410
            lengthRequired = 411
            preconditionFailed = 412
            payloadTooLarge = 413
            urlTooLong = 414
            unsupportedMediaType = 415
            rangeNotSatisfiable = 416
            expectationFailed = 417
            teapot = 418
            methodFailure = 420
            enhanceYourCalm = 420
            misdirectedRequest = 421
            unprocessableEntity = 422
            locked = 423
            failedDependency = 424
            upgradeRequired = 426
            preconditionRequired = 428
            tooManyRequests = 429
            requestHeaderFieldsTooLarge = 431
            loginTimeout = 440
            noResponse = 444
            retryWith = 449
            redirect = 451
            unavailableForLegalReasons = 451
            sslCertificateError = 495
            sslCertificateRequired = 496
            httpRequestSentToSecurePort = 497
            invalidToken = 498
            clientClosedRequest = 499
            tokenRequired = 499
            internalServerError = 500
        End Enum
    End Module
End Namespace
