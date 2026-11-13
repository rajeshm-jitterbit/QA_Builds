pipeline {
  agent { label 'windows' }          // use the label of your Windows agent, or change to: agent any
  triggers { cron('H/15 * * * *') }  // refresh every 15 minutes
  options { timestamps(); disableConcurrentBuilds() }
  environment {
    OCTOPUS_API_KEY = credentials('octopus-api-key')   // Jenkins "Secret text" credential
    // If Python is not on the agent's PATH, set the full path here, e.g.:
    // PY = 'C:\\Users\\<service-account>\\AppData\\Local\\Programs\\Python\\Python312\\python.exe'
    PY = 'python'
  }
  stages {
    stage('Check Python and network') {
      steps {
        bat '%PY% --version'
        powershell 'Resolve-DnsName deploy.jitterbit.us | Out-Null; Write-Host "deploy.jitterbit.us resolves"'
      }
    }
    stage('Install') {
      steps { bat '%PY% -m pip install requests openpyxl' }
    }
    stage('Build tracker') {
      steps { bat '%PY% octopus_version_tracker.py' }
    }
    stage('Publish') {
      steps {
        archiveArtifacts artifacts: 'Harmony_Version_Tracker_Octopus.xlsx, index.html', fingerprint: false
        publishHTML(target: [reportName: 'Harmony Version Tracker', reportDir: '.', reportFiles: 'index.html',
                             keepAll: false, alwaysLinkToLastBuild: true, allowMissing: false])
      }
    }
  }
}
