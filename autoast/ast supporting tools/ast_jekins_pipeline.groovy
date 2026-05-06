pipeline {
  agent { label 'algorab' }

  options {
    // 24-hour max runtime — large submission batches can run long
    timeout(time: 24, unit: 'HOURS')
    // Prevent overlapping nightly runs from racing on the same submission folders in WATCH_DIR
    disableConcurrentBuilds()
    // Do not automatically retry failed builds — BATCH_COMPLETE.json sentinel logic controls reruns
    retry(0)
  }

  // Nightly at 20:59 (server time)
  triggers {
    cron('59 20 * * *')
  }

  environment {
    PYTHON_PATH = '//giswhse.env.gov.bc.ca/whse_np/corp/central_clones/python_geopandas/python.exe'
    SCRIPT_PATH = '//spatialfiles.bcgov/work/srm/nel/Local/Geomatics/Workarea/csostad/GitHub_Repositories/ast_batch_tool/autoast/batch_ast_v3/main_auto_setup_jenkins.py'
    // SDE_CONNECTION_DIR is injected here so geojenk writes bcgw.sde to the objectstore
    // (where it has write access) rather than the personal workarea on spatialfiles (where it does not).
    // This variable is not set in .env so VSCode runs are unaffected — they fall back to connection/ next to the script.
    SDE_CONNECTION_DIR = '\\\\objectstore2.nrs.bcgov\\GSS_Share\\authorizations\\batch_ast_tool\\connection'
    MAIL_LIST   = 'chris.sostad@gov.bc.ca'
    // IMPORTANT: Do NOT set SECRET_FILE here. Let Jenkins env vars drive auth.

  }

  stages {
    stage('Run AutoAST Batch') {
      steps {
        withCredentials([
          usernamePassword(
            credentialsId: 'csostad_bcgw',    // Jenkins credential ID for BCGW Oracle login
            usernameVariable: 'BCGW_USER',    // injected into env; read by database_connection.py
            passwordVariable: 'BCGW_PASS'     // injected into env; read by database_connection.py
          )
        ]) {
          bat """
            "%PYTHON_PATH%" "%SCRIPT_PATH%"
          """
        }
      }
    }
  }

  post {
    failure {
      mail to: env.MAIL_LIST,
           from: 'chris.sostad@gov.bc.ca',
           subject: "AutoAST Batch Failed: ${env.JOB_NAME}",
           body:    "AutoAST Batch Failed - \"${env.JOB_NAME}\" build: ${env.BUILD_NUMBER}\\n\\nView the log at:\\n ${env.BUILD_URL}"
    }
    aborted {
      mail to: env.MAIL_LIST,
           from: 'chris.sostad@gov.bc.ca',
           subject: "AutoAST Batch Aborted: ${env.JOB_NAME}",
           body:    "AutoAST Batch Aborted - \"${env.JOB_NAME}\" build: ${env.BUILD_NUMBER}\\n\\nView the log at:\\n ${env.BUILD_URL}"
    }
    success {
      mail to: env.MAIL_LIST,
           from: 'chris.sostad@gov.bc.ca',
           subject: "AutoAST Batch Succeeded: ${env.JOB_NAME}",
           body:    "AutoAST Batch Succeeded - \"${env.JOB_NAME}\" build: ${env.BUILD_NUMBER}\\n\\nView the log at:\\n ${env.BUILD_URL}"
    }
  }
}
