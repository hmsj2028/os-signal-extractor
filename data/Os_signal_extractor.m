function Os_signal_extractor(swcRootPaths, outputFolder)
%OS_SIGNAL_EXTRACTOR Simulink(.slx) 모델의 CAN 신호 Input/Output 목록을 엑셀로 추출
%
%   사용법 1) 인자 없이 실행 (권장, 사용자 친화적 UI 모드)
%       Os_signal_extractor()
%       -> 1) 라이브러리 폴더가 포함된 상위 폴더 선택 (asw/os/rte 등이 있는 폴더)
%             선택 즉시 하위 전체가 MATLAB 경로에 추가됨
%          2) 상위 폴더 하위를 재귀 검색해서 이름이 정확히 "os"인 폴더를
%             모두 찾음. 1개면 자동 선택, 여러 개면 체크리스트에서 다중 선택
%          3) 선택된 os 폴더(들) 하위의 "os_xxx" 형태 모델 폴더들을 모아
%             체크리스트로 보여주고 다중 선택
%          4) 상위 폴더 하위를 재귀 검색해서 이름이 "rte"인 폴더를 모두 찾고,
%             그 안의 .mat 파일을 전부 base workspace에 로드 (한 번만 수행)
%          5) 선택된 각 os_xxx 폴더 안의 .slx 파일마다: 같은 폴더의
%             "<모델명>_param.m"을 실행한 뒤 모델을 연다 (rte mat 로드가
%             이미 끝난 뒤이므로 param 스크립트가 그 데이터를 참조해도 됨).
%             이름이 'Swc_'로 시작하는 서브시스템(단, 이름에 '_sch'가
%             포함된 스케줄러는 제외)들을 모델을 전부 로드한 뒤에야 알 수
%             있으므로, 모델 로드가 끝날 때까지는 사용자 입력 없이 계속
%             진행됨.
%          (4~5단계는 진행률 표시줄(progress bar)로 진행 상황을 보여줌.
%           소요 시간은 3)단계에서 폴더 선택이 모두 끝난 시점부터 측정)
%          5-1) 모델별 Swc_ 체크리스트 창을 모델 하나당 하나씩 "차례대로"
%             모달로 띄움 (여러 모델을 골랐어도 창이 한꺼번에 뜨지 않고,
%             한 모델을 확인/취소해야 다음 모델 창이 뜸)
%          6) 선택한 모든 모델의 결과를 하나의 엑셀 파일에 모델별 시트로
%             나누어 저장 (Summary 시트에 전체 실행시간 + 모델별 통계를
%             보기 좋게 정리해 포함).
%             파일명은 항상 "result_<날짜시간>.xlsx" (모델 개수/이름과 무관)
%
%   사용법 2) 경로를 직접 지정 (자동화/스크립트용)
%       Os_signal_extractor(swcRootPaths)
%       Os_signal_extractor(swcRootPaths, outputFolder)
%
%   swcRootPaths : 'Swc_'로 시작하는 최상위 서브시스템 경로들의 셀 배열
%   outputFolder : (선택) 엑셀 파일을 저장할 폴더. 미지정 시 현재 폴더(pwd)
%                  (UI 모드에서는 이 스크립트가 있는 폴더의 'results' 하위 폴더가
%                   기본값)
%                  (로그 파일은 이 폴더와 무관하게 항상 이 스크립트가 있는
%                   폴더의 'logs' 하위 폴더에 저장됨)
%
%   각 Swc_ 루트 하위(깊이 제한 없음)에서, 내부에 'PhyVal' 이름의 Inport를
%   가진 마스킹 블록을 "leaf 기능 블록"으로 간주하여 Input/Output을 추출한다.
%     - Input  : PhyVal 포트가 최종적으로 연결된 "SWC 단(Swc_ 서브시스템 자신의
%                경계 Inport)의 입력" 이름. Ground/Constant/게이트가 아닌
%                블록(Goto/From, Data Type Conversion, Gain 등 무엇이든)은
%                전부 신호 원본이 아니라고 보고 이름을 쓰지 않은 채 그 입력까지
%                계속 거슬러 올라간다 - 블록 타입이나 이름을 따지지 않고 항상
%                이렇게 처리하며, Swc_ 서브시스템 자신의 바깥 경계에 닿는
%                순간(=SWC의 입력)에만 멈추고 그 지점의 이름을 사용한다
%                (그 너머 모델 상위의 하네스 신호까지는 따라가지 않음).
%                OR/Switch 같은 게이트 블록을 거치면 게이트로 들어오는 모든
%                신호를 "a, b, c" 형태로 나열하며, Ground/Constant도 그대로
%                블록 이름으로 나열되며 Note에 부연)
%     - ErrFlagInput : 같은 블록에 'ErrFlag' 포트가 있으면 동일한 방식으로 추출
%     - Output : leaf 블록 자신의 이름
%     - Note   : 그라운드는 "ground(0)", 게이트 다중입력은 "다중입력 N건",
%                상수 입력은 "constant(값)" 형태로 표기 (여러 개면 쉼표로 나열)
%   버전별로 동일한 Output이 중복되면(같은 모델 내에서) 먼저 발견된 것만
%   남기고 이후 것은 버리며, 만약 Input 값이 서로 다르면 그 사실을 로그에
%   WARN으로 남긴다 (중복 자체는 모델별로만 비교하며, 다른 모델끼리는
%   서로 중복으로 취급하지 않음).
%   실행마다 결과 엑셀(모델별 Data 시트 + Summary 시트)과 디버그 로그 텍스트
%   파일을 생성한다. 로그 파일에는 성공적으로 처리된 내용은 남기지 않고
%   경고(WARN)/오류(ERROR)와 시작-종료 구분선만 남겨 특이사항만 바로 보이게 한다.

    % ---------------- 설정값 (필요 시 이 부분만 조정) ----------------
    CFG.PhyValName          = 'PhyVal';
    CFG.ErrFlagName         = 'ErrFlag';
    CFG.GateBlockTypes      = {'Logic', 'Switch', 'MultiPortSwitch'}; % 다중 입력을 모두 나열할 블록 타입
    CFG.MaxTraceDepth       = 30;   % Goto/게이트 추적 시 순환 참조 방지용 최대 깊이
    CFG.SwcNamePattern      = '^Swc_';  % 대소문자 무시하고 이 패턴으로 시작하는 서브시스템을 루트로 인식
    CFG.SchedulerNamePattern = '_sch'; % 대소문자 무시하고 이 문자열을 포함하면 스케줄러로 간주해 제외
    % ------------------------------------------------------------------

    uiMode = (nargin == 0);
    logFile = -1;

    if uiMode
        parentFolder = getParentFolder();
        if isempty(parentFolder)
            return; % 사용자가 취소함
        end
        addpath(genpath(parentFolder));

        osFolders = selectOsFolders(parentFolder);
        if isempty(osFolders)
            uiwait(warndlg('"os" 폴더를 찾지 못했습니다.', '알림'));
            return;
        end

        modelFolders = selectModelFolders(osFolders, parentFolder);
        if isempty(modelFolders)
            return; % 사용자가 취소함
        end

        startTime = now; % 폴더 선택이 모두 끝난 시점부터 소요 시간을 측정 (사용자가 고민하며 고르는 대기 시간은 제외)

        logFile = openLogFile();

        loadAllRteMatFiles(parentFolder, logFile);

        % --- 1단계: 시간이 걸리는 모델 로드 + Swc_ 탐색을 먼저 전부 끝낸다 ---
        % (대화상자 없이 순차 진행 - 모델을 여러 개 선택했을 때 "로드 대기 ->
        %  선택 -> 로드 대기 -> 선택"이 반복되며 기다리는 게 불편하다는 피드백 반영)
        pendingSelections = struct('modelName', {}, 'discovered', {}, 'folderBaseName', {});

        totalSlxFiles = 0;
        for i = 1:numel(modelFolders)
            totalSlxFiles = totalSlxFiles + numel(dir(fullfile(modelFolders{i}, '*.slx')));
        end
        hProgress = safeWaitbarCreate(0, '모델 로드 준비 중...', 'OS 신호 추출기 - 진행 상황');
        processedSlx = 0;

        for i = 1:numel(modelFolders)
            modelFolder = modelFolders{i};
            [~, folderBaseName] = fileparts(modelFolder);
            slxFiles = dir(fullfile(modelFolder, '*.slx'));
            if isempty(slxFiles)
                logMsg(logFile, 'WARN', '.slx 파일을 찾지 못해 건너뜁니다: %s', modelFolder);
                continue;
            end

            for j = 1:numel(slxFiles)
                fullPath = fullfile(modelFolder, slxFiles(j).name);
                [~, modelNameForLoad, ~] = fileparts(fullPath);
                safeWaitbarUpdate(hProgress, processedSlx / max(totalSlxFiles, 1), ...
                    sprintf('모델 로드 중 (%d/%d)\n%s', processedSlx + 1, totalSlxFiles, modelNameForLoad));

                loadModelParamScript(modelFolder, modelNameForLoad, logFile);

                modelName = loadModelFromFile(fullPath);
                processedSlx = processedSlx + 1;
                if isempty(modelName)
                    continue;
                end

                discovered = discoverSwcRoots(modelName, CFG);
                if isempty(discovered)
                    logMsg(logFile, 'WARN', '"%s" 모델에서 "%s" 로 시작하는 서브시스템을 찾지 못함', ...
                        modelName, CFG.SwcNamePattern);
                    continue;
                end

                pendingSelections(end + 1) = struct('modelName', modelName, ...
                    'discovered', {discovered}, 'folderBaseName', folderBaseName); %#ok<AGROW>
            end
        end
        safeWaitbarClose(hProgress);

        if isempty(pendingSelections)
            uiwait(warndlg('선택한 모델들에서 "Swc_" 서브시스템을 찾지 못했습니다. 로그 파일을 확인해 주세요.', '알림'));
            if logFile > 0
                fclose(logFile);
            end
            return;
        end

        % --- 2단계: 로드가 모두 끝난 뒤, 모델별 선택 창을 하나씩 차례대로 띄운다 ---
        % (모델을 여러 개 골랐을 때 창이 한꺼번에 쫙 뜨면 헷갈리므로, 모델별로
        %  선택 창을 하나씩 모달로 띄워 확인/취소해야 다음 모델 창이 뜨게 함)
        swcRootPaths = selectSubsystemsUIMulti(pendingSelections);

        if isempty(swcRootPaths)
            if logFile > 0
                fclose(logFile);
            end
            return;
        end
        % 이 파일은 도구 폴더의 "data" 하위에 있으므로, 한 단계 위(도구
        % 폴더 자체)를 기준으로 results 폴더를 둔다 - data 폴더 안에
        % 결과물이 섞여 들어가지 않게 하기 위함.
        outputFolder = fullfile(fileparts(fileparts(mfilename('fullpath'))), 'results');
    else
        startTime = now; % 경로 지정 모드는 폴더 선택 UI가 없으므로 호출 시점부터 측정
        if ischar(swcRootPaths)
            swcRootPaths = {swcRootPaths};
        end
        if nargin < 2 || isempty(outputFolder)
            outputFolder = pwd;
        end
    end

    result = runExtraction(swcRootPaths, outputFolder, CFG, logFile, startTime);

    if uiMode
        showSummaryDialog(result);
    end
end

% ====================================================================
% UI 관련 함수
% ====================================================================
function h = safeWaitbarCreate(frac, msg, titleStr)
% 진행 상황 표시용 waitbar를 만든다. 디스플레이가 없는 등의 이유로 실패해도
% 전체 추출 작업이 중단되지 않도록 항상 방어한다.
    h = [];
    try
        h = waitbar(frac, escapeTexUnderscore(msg), 'Name', titleStr);
    catch
    end
end

function safeWaitbarUpdate(h, frac, msg)
    if isempty(h) || ~ishandle(h)
        return;
    end
    try
        waitbar(min(max(frac, 0), 1), h, escapeTexUnderscore(msg));
    catch
    end
end

function s = escapeTexUnderscore(s)
% waitbar 메시지는 기본적으로 TeX 인터프리터로 해석되어 "_"가 있으면
% 아래첨자로 표시되면서 밑줄이 사라지고 뒤 글자가 아래로 내려가 보인다
% (예: "Swc_..." -> "Sw" + 아래첨자 "c..."). Swc_ 서브시스템/모델 이름에
% "_"가 흔히 들어가므로, 문자 그대로 보이도록 이스케이프한다.
    s = strrep(s, '_', '\_');
end

function safeWaitbarClose(h)
    if isempty(h) || ~ishandle(h)
        return;
    end
    try
        close(h);
    catch
    end
end

function parentFolder = getParentFolder()
% 라이브러리 폴더(및 asw/os/rte 등)가 들어있는 상위 폴더를 매번 새로 선택받는다.
    parentFolder = '';

    selected = uigetdir(pwd, '라이브러리 폴더가 포함된 상위 폴더를 선택하세요 (asw/os/rte 등이 있는 폴더)');
    if isequal(selected, 0)
        parentFolder = '';
        return;
    end
    parentFolder = selected;
end

function osFolders = selectOsFolders(parentFolder)
% parentFolder 하위를 재귀 검색해서 이름이 정확히 "os"인 폴더를 모두 찾는다.
% 1개면 그대로 사용, 여러 개면 체크리스트에서 다중 선택하게 한다
% (서로 다른 모듈 트리 밑에 os 폴더가 여러 개 있을 수 있으므로).
    candidates = findFoldersByName(parentFolder, 'os');

    if isempty(candidates)
        osFolders = {};
        return;
    end
    if numel(candidates) == 1
        osFolders = candidates;
        return;
    end

    displayNames = cell(size(candidates));
    for i = 1:numel(candidates)
        displayNames{i} = relativeToParent(candidates{i}, parentFolder);
    end

    [sel, ok] = listdlg( ...
        'PromptString', '"os" 폴더가 여러 개 발견되었습니다. 사용할 os 폴더를 선택하세요 (다중 선택 가능)', ...
        'ListString', displayNames, ...
        'SelectionMode', 'multiple', ...
        'InitialValue', 1:numel(candidates), ...
        'ListSize', [640, 320], ...
        'Name', 'os 폴더 선택');

    if ~ok
        osFolders = {};
        return;
    end
    osFolders = candidates(sel);
end

function modelFolders = selectModelFolders(osFolders, parentFolder)
% 선택된 os 폴더(들) 바로 하위에서, 내부에 .slx 파일이 있는 하위 폴더
% (os_xxx 형태)들을 모아 체크리스트로 다중 선택하게 한다.
    candidates = {};
    for i = 1:numel(osFolders)
        items = dir(osFolders{i});
        for j = 1:numel(items)
            if items(j).isdir && ~any(strcmp(items(j).name, {'.', '..'}))
                sub = fullfile(osFolders{i}, items(j).name);
                if ~isempty(dir(fullfile(sub, '*.slx')))
                    candidates{end + 1} = sub; %#ok<AGROW>
                end
            end
        end
    end

    if isempty(candidates)
        modelFolders = {};
        uiwait(warndlg('.slx 파일이 들어있는 모델 폴더(os_xxx)를 찾지 못했습니다.', '알림'));
        return;
    end

    displayNames = cell(size(candidates));
    for i = 1:numel(candidates)
        displayNames{i} = relativeToParent(candidates{i}, parentFolder);
    end

    [sel, ok] = listdlg( ...
        'PromptString', '추출할 모델 폴더(os_xxx)를 선택하세요 (다중 선택 가능)', ...
        'ListString', displayNames, ...
        'SelectionMode', 'multiple', ...
        'InitialValue', 1:numel(candidates), ...
        'ListSize', [640, 320], ...
        'Name', '모델 폴더 선택');

    if ~ok
        modelFolders = {};
        return;
    end
    modelFolders = candidates(sel);
end

function found = findFoldersByName(rootFolder, targetName)
% rootFolder 하위를 재귀적으로 훑어서 이름이 targetName과 일치(대소문자 무시)하는
% 폴더 경로를 모두 찾는다.
    found = {};
    stack = {rootFolder};
    while ~isempty(stack)
        cur = stack{end};
        stack(end) = [];

        items = dir(cur);
        for i = 1:numel(items)
            if ~items(i).isdir || any(strcmp(items(i).name, {'.', '..'}))
                continue;
            end
            sub = fullfile(cur, items(i).name);
            if strcmpi(items(i).name, targetName)
                found{end + 1} = sub; %#ok<AGROW>
            end
            stack{end + 1} = sub; %#ok<AGROW>
        end
    end
end

function rel = relativeToParent(fullPath, parentFolder)
    rel = fullPath;
    if startsWith(rel, parentFolder)
        rel = rel(numel(parentFolder) + 1:end);
        rel = regexprep(rel, '^[\\/]+', '');
    end
end

function loadAllRteMatFiles(parentFolder, logFile)
% parentFolder 하위를 재귀 검색해서 이름이 정확히 "rte"인 폴더를 모두 찾고,
% 그 안의 .mat 파일을 전부 base workspace에 로드한다 (os 폴더가 여러 모듈
% 트리에 흩어져 있는 것처럼 rte도 흩어져 있을 수 있으므로 전체를 훑는다).
% param 스크립트들이 이 데이터를 참조하는 경우가 많으므로 모델 로드보다 먼저,
% 한 번만 실행한다.
    rteFolders = findFoldersByName(parentFolder, 'rte');
    if isempty(rteFolders)
        logMsg(logFile, 'WARN', '"rte" 폴더를 찾지 못함: %s', parentFolder);
        return;
    end

    allMatFiles = {};
    for f = 1:numel(rteFolders)
        matFiles = dir(fullfile(rteFolders{f}, '*.mat'));
        if isempty(matFiles)
            logMsg(logFile, 'WARN', 'rte 폴더에 .mat 파일이 없습니다: %s', rteFolders{f});
        end
        for i = 1:numel(matFiles)
            allMatFiles{end + 1} = fullfile(matFiles(i).folder, matFiles(i).name); %#ok<AGROW>
        end
    end

    total = numel(allMatFiles);
    hProgress = safeWaitbarCreate(0, '.mat 파일 로드 준비 중...', 'OS 신호 추출기 - 진행 상황');
    for i = 1:total
        matPath = allMatFiles{i};
        safeWaitbarUpdate(hProgress, (i - 1) / max(total, 1), ...
            sprintf('.mat 파일 로드 중 (%d/%d)\n%s', i, total, matPath));
        try
            varsBefore = evalin('base', 'who');
            evalin('base', sprintf('load(''%s'');', strrep(matPath, '''', '''''')));
            varsAfter = evalin('base', 'who');
            logMsg(logFile, 'INFO', '"%s" 로드 완료 (변수 %d개 추가됨)', matPath, numel(varsAfter) - numel(varsBefore));
        catch ME
            logMsg(logFile, 'ERROR', '"%s" 로드 실패: %s', matPath, ME.message);
        end
    end
    safeWaitbarClose(hProgress);
end

function loadModelParamScript(modelFolder, modelName, logFile)
% os_xxx 폴더의 "<모델명>_param.m" 스크립트를 base workspace에서 실행한다.
% (rte의 mat 데이터를 참조하는 경우가 많으므로, loadAllRteMatFiles가 먼저
%  끝난 뒤에 호출되어야 한다)
    paramFile = fullfile(modelFolder, [modelName '_param.m']);
    if isfile(paramFile)
        try
            evalin('base', sprintf('run(''%s'');', strrep(paramFile, '''', '''''')));
            logMsg(logFile, 'INFO', '"%s" 실행 완료', paramFile);
        catch ME
            logMsg(logFile, 'ERROR', '"%s" 실행 실패: %s', paramFile, ME.message);
        end
    else
        logMsg(logFile, 'WARN', '_param.m 파일을 찾지 못함: %s', paramFile);
    end
end

function modelName = loadModelFromFile(fullPath)
    [~, modelName, ~] = fileparts(fullPath);
    try
        if ~bdIsLoaded(modelName)
            load_system(fullPath);
        end
    catch ME
        uiwait(errordlg(escapeTexUnderscore(sprintf('모델 로드 실패: %s\n%s', fullPath, ME.message)), '오류'));
        modelName = '';
    end
end

function swcRootPaths = selectSubsystemsUIMulti(pendingSelections)
% 모델별로 Swc_ 서브시스템 선택 창을 띄운다. 여러 모델을 한꺼번에 골랐어도
% 창을 전부 동시에 띄우지 않고, 모델 하나당 하나씩 차례대로(모달) 띄워서
% 사용자가 순서대로 확인/취소하게 한다.
    n = numel(pendingSelections);
    resultsCell = cell(n, 1);

    for i = 1:n
        resultsCell{i} = selectSubsystemsUISingle(pendingSelections(i));
    end

    swcRootPaths = {};
    for i = 1:n
        swcRootPaths = [swcRootPaths, resultsCell{i}]; %#ok<AGROW>
    end
end

function selected = selectSubsystemsUISingle(pending)
% 한 모델에 대해 Swc_ 서브시스템 선택 창을 모달로 띄우고, 사용자가 확인/취소를
% 누르거나 창을 닫을 때까지 기다렸다가 선택 결과를 반환한다.
    modelName = pending.modelName;
    discovered = pending.discovered;
    folderBaseName = pending.folderBaseName;

    prefix = [modelName '/'];
    displayNames = cell(size(discovered));
    for k = 1:numel(discovered)
        d = discovered{k};
        if startsWith(d, prefix)
            displayNames{k} = d(numel(prefix) + 1:end);
        else
            displayNames{k} = d;
        end
    end

    screenSize = get(0, 'ScreenSize');
    winWidth = 420;
    winHeight = 380;
    posX = max(20, (screenSize(3) - winWidth) / 2);
    posY = max(20, (screenSize(4) - winHeight) / 2);

    selected = {};

    f = figure('Name', sprintf('[%s] %s - 대상 서브시스템 선택', folderBaseName, modelName), ...
        'NumberTitle', 'off', 'MenuBar', 'none', 'ToolBar', 'none', ...
        'Position', [posX, posY, winWidth, winHeight], ...
        'Resize', 'off', 'WindowStyle', 'modal', 'CloseRequestFcn', @(src, evt) onCancel());

    uicontrol(f, 'Style', 'text', 'Units', 'normalized', ...
        'Position', [0.02, 0.90, 0.96, 0.08], 'HorizontalAlignment', 'left', ...
        'String', sprintf('"%s" 모델 - 추출할 Swc_ 서브시스템 (기본: 전체 선택)', modelName));

    lb = uicontrol(f, 'Style', 'listbox', 'Units', 'normalized', ...
        'Position', [0.02, 0.15, 0.96, 0.75], ...
        'String', displayNames, 'Max', 2, 'Min', 0, ...
        'Value', 1:numel(displayNames));

    uicontrol(f, 'Style', 'pushbutton', 'Units', 'normalized', ...
        'Position', [0.55, 0.02, 0.42, 0.09], 'String', '확인', ...
        'Callback', @(src, evt) onConfirm());
    uicontrol(f, 'Style', 'pushbutton', 'Units', 'normalized', ...
        'Position', [0.02, 0.02, 0.42, 0.09], 'String', '취소', ...
        'Callback', @(src, evt) onCancel());

    uiwait(f);

    function onConfirm()
        sel = get(lb, 'Value');
        selected = discovered(sel);
        delete(f);
    end

    function onCancel()
        selected = {};
        if ishandle(f)
            delete(f);
        end
    end
end

function showSummaryDialog(result)
    if isempty(result.generatedFiles)
        uiwait(warndlg('추출된 결과가 없습니다. 로그 파일을 확인해 주세요.', '완료'));
        return;
    end

    msg = sprintf('%s\n\n결과 파일:\n%s\n\n로그 파일:\n%s', ...
        strjoin(result.reportLines, newline), strjoin(result.generatedFiles, newline), result.logFilePath);
    uiwait(msgbox(escapeTexUnderscore(msg), '완료'));
end

% ====================================================================
% 모델 탐색 함수
% ====================================================================
function roots = discoverSwcRoots(modelName, CFG)
% 모델 전체에서 이름이 CFG.SwcNamePattern으로 시작하는 서브시스템을 찾는다
% (계층 깊이와 무관하게 동작 - 모델마다 구조가 달라도 됨)
% FollowLinks를 켜서 라이브러리에서 가져온(링크된) 서브시스템도 탐색 대상에 포함시키고,
% 이름에 CFG.SchedulerNamePattern이 포함된 스케줄러 서브시스템은 결과에서 제외한다.
    roots = {};
    allSubsystems = find_system(modelName, 'FollowLinks', 'on', 'LookUnderMasks', 'all', 'BlockType', 'SubSystem');
    for i = 1:numel(allSubsystems)
        sys = allSubsystems{i};
        nm = get_param(sys, 'Name');
        if isempty(regexpi(nm, CFG.SwcNamePattern, 'once'))
            continue;
        end
        if ~isempty(regexpi(nm, CFG.SchedulerNamePattern, 'once'))
            continue; % 스케줄러는 leaf 블록 탐색 대상/선택 목록에서 제외
        end
        roots{end + 1} = sys; %#ok<AGROW>
    end
end

function leafBlocks = findLeafBlocks(swcPath, CFG, logFile)
% swcPath 하위 전체(깊이 제한 없음)에서 마스킹된 블록 중, 내부에
% PhyVal 이름의 Inport를 가진 것만 "leaf 기능 블록"으로 필터링
% FollowLinks를 켜서 라이브러리 링크로 가져온 블록 내부까지 탐색한다.
    candidates = find_system(swcPath, 'FollowLinks', 'on', 'LookUnderMasks', 'all', 'Mask', 'on');
    candidates(strcmp(candidates, swcPath)) = [];

    leafBlocks = {};
    for i = 1:numel(candidates)
        if hasNamedInnerInport(candidates{i}, CFG.PhyValName)
            leafBlocks{end + 1} = candidates{i}; %#ok<AGROW>
        end
    end

    if isempty(leafBlocks)
        logMsg(logFile, 'WARN', '%s 하위에서 %s 입력을 가진 마스킹 블록을 찾지 못함', swcPath, CFG.PhyValName);
    end
end

% ====================================================================
% 추출 로직 (경로 지정 모드 / UI 모드 공통)
% ====================================================================
function result = runExtraction(swcRootPaths, outputFolder, CFG, logFile, startTime)
    if nargin < 4 || isempty(logFile) || logFile < 0
        logFile = openLogFile();
    end
    if nargin < 5 || isempty(startTime)
        startTime = now;
    end

    logMsg(logFile, 'STAGE', '===== 추출 시작 (대상 %d개) =====', numel(swcRootPaths));

    resultsByModel   = containers.Map('KeyType', 'char', 'ValueType', 'any');
    seenOutputByModel = containers.Map('KeyType', 'char', 'ValueType', 'any');
    statsByModel     = containers.Map('KeyType', 'char', 'ValueType', 'any');

    totalRoots = numel(swcRootPaths);
    hProgress = safeWaitbarCreate(0, '신호 추출 준비 중...', 'OS 신호 추출기 - 진행 상황');

    for k = 1:numel(swcRootPaths)
        swcPath = swcRootPaths{k};
        modelName = extractModelName(swcPath);
        safeWaitbarUpdate(hProgress, (k - 1) / max(totalRoots, 1), ...
            sprintf('신호 추출 중 (%d/%d)\n%s', k, totalRoots, swcPath));

        if ~ensureModelLoaded(modelName, logFile)
            continue;
        end
        if ~isValidPath(swcPath)
            logMsg(logFile, 'WARN', '경로를 찾을 수 없어 건너뜁니다: %s', swcPath);
            continue;
        end

        if ~isKey(resultsByModel, modelName)
            resultsByModel(modelName) = cell(0, 5); % SubsystemPath, Input, ErrFlagInput, Output, Note
            seenOutputByModel(modelName) = containers.Map('KeyType', 'char', 'ValueType', 'any');
            statsByModel(modelName) = struct('ground', 0, 'gateMulti', 0, 'dupRemoved', 0, 'unresolved', 0);
        end

        logMsg(logFile, 'INFO', '--- Swc 루트 처리 시작: %s ---', swcPath);
        leafBlocks = findLeafBlocks(swcPath, CFG, logFile);
        logMsg(logFile, 'INFO', '%s 에서 leaf 블록 %d개 발견', swcPath, numel(leafBlocks));

        rows = resultsByModel(modelName);
        seenOutputs = seenOutputByModel(modelName);
        modelStats = statsByModel(modelName);

        for i = 1:numel(leafBlocks)
            blockPath = leafBlocks{i};
            try
                [row, flags] = extractLeafRow(blockPath, swcPath, CFG, logFile);
            catch ME
                logMsg(logFile, 'ERROR', '블록 처리 중 오류 (%s): %s', blockPath, ME.message);
                continue;
            end

            outputName = row{4};

            if isKey(seenOutputs, outputName)
                prevRow = seenOutputs(outputName);
                if ~isequal(prevRow{2}, row{2})
                    logMsg(logFile, 'WARN', ...
                        '중복 Output "%s" - Input 값 불일치 (기존:"%s" / 신규:"%s", 신규경로:%s) -> 기존 값 유지', ...
                        outputName, prevRow{2}, row{2}, blockPath);
                end
                modelStats.dupRemoved = modelStats.dupRemoved + 1;
                continue;
            end

            seenOutputs(outputName) = row;
            rows(end + 1, :) = row; %#ok<AGROW>

            if flags.isGround,     modelStats.ground     = modelStats.ground + 1;     end
            if flags.isGateMulti,  modelStats.gateMulti  = modelStats.gateMulti + 1;  end
            if flags.isUnresolved, modelStats.unresolved = modelStats.unresolved + 1; end
        end

        resultsByModel(modelName) = rows;
        seenOutputByModel(modelName) = seenOutputs;
        statsByModel(modelName) = modelStats;
    end

    generatedFiles = {};
    reportLines = {};
    modelNames = keys(resultsByModel);

    % 신호 추출 자체는 끝났지만, 엑셀 파일 생성(Excel COM으로 서식까지
    % 입히는 과정)은 데이터가 많으면 꽤 걸릴 수 있다. 여기서 진행률 표시줄을
    % 바로 닫아버리면 "완료" 안내창이 뜨기 전까지 화면에 아무 것도 안 떠서
    % 멈춘 것처럼 보인다는 피드백이 있어, 진행률 표시줄을 가득 찬 상태로
    % 유지한 채 엑셀 생성이 끝난 뒤에야 닫는다.
    if ~isempty(modelNames)
        safeWaitbarUpdate(hProgress, 1, '결과 엑셀 파일 생성 중...');
        outFile = writeCombinedExcel(resultsByModel, statsByModel, outputFolder, logFile, startTime);
        generatedFiles{end + 1} = outFile;
    end
    safeWaitbarClose(hProgress);

    for i = 1:numel(modelNames)
        modelName = modelNames{i};
        rows = resultsByModel(modelName);
        modelStats = statsByModel(modelName);
        reportLines{end + 1} = sprintf('%s : %d건 (그라운드 %d, 다중입력 %d, 중복제거 %d, 미연결 %d)', ...
            modelName, size(rows, 1), modelStats.ground, modelStats.gateMulti, ...
            modelStats.dupRemoved, modelStats.unresolved); %#ok<AGROW>
    end

    logMsg(logFile, 'STAGE', '===== 추출 종료 =====');
    logFilePath = '';
    if logFile > 0
        logFilePath = fopen(logFile);
        fclose(logFile);
    end

    result = struct('generatedFiles', {generatedFiles}, 'reportLines', {reportLines}, 'logFilePath', logFilePath);
end

function modelName = extractModelName(targetPath)
    tokens = strsplit(targetPath, '/');
    modelName = tokens{1};
end

function ok = ensureModelLoaded(modelName, logFile)
    ok = true;
    if ~bdIsLoaded(modelName)
        try
            load_system(modelName);
        catch ME
            logMsg(logFile, 'ERROR', '모델을 로드하지 못했습니다 (%s): %s', modelName, ME.message);
            ok = false;
        end
    end
end

function ok = isValidPath(targetPath)
    ok = true;
    try
        get_param(targetPath, 'Handle');
    catch
        ok = false;
    end
end

% ====================================================================
% leaf 블록 1개 처리
% ====================================================================
function [row, flags] = extractLeafRow(blockPath, swcPath, CFG, logFile)
    outputName = get_param(blockPath, 'Name');

    [inputStr, flags] = resolvePortInput(blockPath, CFG.PhyValName, swcPath, CFG, logFile);

    errFlagStr = '';
    if hasNamedInnerInport(blockPath, CFG.ErrFlagName)
        [errFlagStr, ~] = resolvePortInput(blockPath, CFG.ErrFlagName, swcPath, CFG, logFile);
    end

    noteParts = {};
    if flags.isGround
        noteParts{end + 1} = 'ground(0)';
    end
    if flags.isGateMulti
        noteParts{end + 1} = sprintf('다중입력 %d건', flags.count);
    end
    if ~isempty(flags.constantValues)
        constStrs = cellfun(@(v) sprintf('constant(%s)', v), flags.constantValues, 'UniformOutput', false);
        noteParts{end + 1} = strjoin(constStrs, ', ');
    end
    if flags.isUnresolved
        noteParts{end + 1} = '미연결';
    end
    note = strjoin(noteParts, '; ');

    row = {swcPath, inputStr, errFlagStr, outputName, note};
end

function has = hasNamedInnerInport(blockPath, portName)
    has = ~isempty(findInnerInportByName(blockPath, portName));
end

function blk = findInnerInportByName(blockPath, portName)
% 마스크 내부에서 이름이 portName과 일치(대소문자 무시)하거나, 그게 없으면
% portName을 포함하는 이름의 Inport 블록을 찾는다
% FollowLinks를 켜서 blockPath 자체가 라이브러리 링크 블록이어도 내부를 탐색할 수 있게 한다.
    blk = '';
    innerInports = find_system(blockPath, 'FollowLinks', 'on', 'LookUnderMasks', 'all', 'SearchDepth', 1, 'BlockType', 'Inport');

    for j = 1:numel(innerInports)
        if strcmpi(get_param(innerInports{j}, 'Name'), portName)
            blk = innerInports{j};
            return;
        end
    end
    for j = 1:numel(innerInports)
        if contains(lower(get_param(innerInports{j}, 'Name')), lower(portName))
            blk = innerInports{j};
            return;
        end
    end
end

function outerHandle = mapInnerInportToOuterHandle(innerBlockPath, portHandles)
    outerHandle = [];
    portNum = str2double(get_param(innerBlockPath, 'Port'));
    if ~isnan(portNum) && portNum >= 1 && portNum <= numel(portHandles.Inport)
        outerHandle = portHandles.Inport(portNum);
    end
end

function [str, flags] = resolvePortInput(blockPath, portName, swcPath, CFG, logFile)
    flags = struct('isGround', false, 'isGateMulti', false, 'isUnresolved', false, 'count', 0, ...
        'constantValues', {{}});
    str = '';

    portHandles = get_param(blockPath, 'PortHandles');
    if isempty(portHandles.Inport)
        flags.isUnresolved = true;
        return;
    end

    innerBlk = findInnerInportByName(blockPath, portName);
    outerHandle = [];
    if ~isempty(innerBlk)
        outerHandle = mapInnerInportToOuterHandle(innerBlk, portHandles);
    end
    if isempty(outerHandle)
        outerHandle = portHandles.Inport(1); % 이름으로 못 찾으면 첫 번째 Inport로 대체
    end

    [names, constVals] = traceSourceNames(outerHandle, swcPath, 0, CFG, logFile);

    if isempty(names)
        flags.isUnresolved = true;
        return;
    end

    if all(strcmp(names, '__GROUND__'))
        flags.isGround = true;
        return;
    end

    displayNames = strrep(names, '__GROUND__', 'GND');
    if numel(displayNames) > 1
        flags.isGateMulti = true;
        flags.count = numel(displayNames);
    end
    str = strjoin(displayNames, ', ');

    flags.constantValues = constVals(~cellfun(@isempty, constVals));
end

function [names, constVals] = traceSourceNames(portHandle, swcPath, depth, CFG, logFile)
% 포트에 연결된 라인을 따라 원본 소스 이름(들)을 찾는다.
% swcPath : 지금 추출 중인 Swc_ 서브시스템의 경로. 추적이 이 경계를 벗어나려는
%           순간(서브시스템 Inport 또는 scoped/global Goto를 통해) 멈추고,
%           그 경계 지점의 이름을 "SWC 단의 입력"으로 보고한다 - 그 너머
%           (모델의 더 상위 하네스 신호 등)까지는 따라가지 않는다.
%   - Ground 소스면 {'__GROUND__'} 마커를 반환
%   - Constant 소스면 블록 이름을 반환하되, constVals에 그 값을 함께 담아
%     Note에 "constant(값)"으로 표기할 수 있게 함
%   - From 이면 같은 GotoTag의 Goto까지 거슬러 올라가 재귀 추적 (단, 그 Goto가
%     swcPath 밖에 있으면 거기서 멈춤)
%   - Inport 소스면(=서브시스템 경계) swcPath 자신의 경계이거나 모델
%     최상위 포트면 거기서 멈추고, 그 외 중간 서브시스템 경계라면 한 단계
%     위 서브시스템의 대응 포트로 계속 거슬러 올라감
%   - CFG.GateBlockTypes(OR/Switch 등)면 그 블록의 모든 Inport를 재귀 추적해 나열
%   - 그 외(Data Type Conversion, Gain 등 무엇이든)에는 신호 원본이 아니라고
%     보고 이름을 쓰지 않은 채 그 블록의 입력까지 계속 거슬러 올라감
%   각 단계는 개별적으로 방어(try/catch)되어 있어, 한 곳에서 에러가 나도
%   전체 행이 통째로 사라지지 않고 [오류:...] 표시와 함께 로그에 상세히 남는다.
    names = {};
    constVals = {};
    if depth > CFG.MaxTraceDepth
        logMsg(logFile, 'WARN', '추적 깊이(%d) 초과 - 순환 참조 의심, 추적 중단', CFG.MaxTraceDepth);
        return;
    end

    try
        lineHandle = get_param(portHandle, 'Line');
    catch ME
        logMsg(logFile, 'ERROR', '포트 정보 조회 실패: %s', ME.message);
        names = {'[오류:포트조회실패]'};
        constVals = {''};
        return;
    end
    if lineHandle == -1
        return; % 연결된 라인 없음
    end

    try
        srcBlockHandle = get_param(lineHandle, 'SrcBlockHandle');
    catch ME
        logMsg(logFile, 'ERROR', '라인의 소스 블록 조회 실패: %s', ME.message);
        names = {'[오류:소스조회실패]'};
        constVals = {''};
        return;
    end
    if srcBlockHandle == -1
        return;
    end

    try
        blockType = get_param(srcBlockHandle, 'BlockType');
        blockPath = getfullname(srcBlockHandle);
    catch ME
        logMsg(logFile, 'ERROR', '블록 정보 조회 실패 (handle=%d): %s', srcBlockHandle, ME.message);
        names = {'[오류:블록정보조회실패]'};
        constVals = {''};
        return;
    end

    switch blockType
        case 'Ground'
            names = {'__GROUND__'};
            constVals = {''};

        case 'Constant'
            names = {get_param(srcBlockHandle, 'Name')};
            constVals = {get_param(srcBlockHandle, 'Value')};

        case 'From'
            [names, constVals] = resolveFrom(srcBlockHandle, blockPath, swcPath, depth, CFG, logFile);

        case 'Inport'
            [names, constVals] = crossSubsystemBoundary(srcBlockHandle, blockPath, swcPath, depth, CFG, logFile);

        otherwise
            if any(strcmp(blockType, CFG.GateBlockTypes))
                [names, constVals] = resolveGate(srcBlockHandle, blockType, swcPath, depth, CFG, logFile);
            else
                % Ground/Constant/From/Inport/게이트가 아닌 블록(Data Type
                % Conversion, Gain, 그 외 무엇이든)은 우리가 원하는 최종
                % 결과("SWC 단의 인포트")가 아니므로, 블록 이름을 쓰지 않고
                % 그 입력까지 계속 거슬러 올라간다. 블록 타입이나 이름을
                % 따지지 않고 항상 이렇게 처리한다.
                [names, constVals] = resolvePassthrough(srcBlockHandle, swcPath, depth, CFG, logFile);
            end
    end
end

function [names, constVals] = resolvePassthrough(blockHandle, swcPath, depth, CFG, logFile)
% Ground/Constant/From/Inport/게이트가 아닌 블록은 그 자체가 최종 목적지
% (SWC 단의 인포트)가 아니므로, 이 블록 이름을 쓰지 않고 이 블록의 입력까지
% 계속 거슬러 올라간다.
    try
        portHandles = get_param(blockHandle, 'PortHandles');
        if isempty(portHandles.Inport)
            names = {get_param(blockHandle, 'Name')};
            constVals = {''};
            return;
        end
        [names, constVals] = traceSourceNames(portHandles.Inport(1), swcPath, depth + 1, CFG, logFile);
        if isempty(names)
            names = {get_param(blockHandle, 'Name')};
            constVals = {''};
        end
    catch ME
        logMsg(logFile, 'ERROR', '통과성 블록 "%s" 처리 중 오류: %s', get_param(blockHandle, 'Name'), ME.message);
        names = {'[오류:통과블록처리실패]'};
        constVals = {''};
    end
end

function [names, constVals] = crossSubsystemBoundary(inportHandle, inportPath, swcPath, depth, CFG, logFile)
% Inport는 서브시스템의 "경계 포트"일 뿐, 그 자체가 신호 원본이 아닐 수 있다.
% 단, 이 Inport의 부모가 지금 추출 중인 Swc_ 서브시스템(swcPath) 자신이거나
% 모델 최상위라면, 그게 바로 우리가 보고하고 싶은 "SWC 단의 입력"이므로
% 여기서 멈춘다. 그 외(=swcPath 내부의 더 깊은 중간 서브시스템 경계)라면
% 한 단계 위 서브시스템의 대응 외부 포트로 넘어가서 계속 추적한다.
    constVals = {};
    try
        model = bdroot(inportPath);
        parentPath = get_param(inportHandle, 'Parent');

        if strcmp(parentPath, swcPath) || strcmp(parentPath, model)
            names = {get_param(inportHandle, 'Name')};
            constVals = {''};
            return;
        end

        portNum = str2double(get_param(inportHandle, 'Port'));
        if isnan(portNum)
            logMsg(logFile, 'WARN', 'Inport "%s"의 Port 번호를 읽지 못해 이름으로 대체', inportPath);
            names = {get_param(inportHandle, 'Name')};
            constVals = {''};
            return;
        end

        parentPortHandles = get_param(parentPath, 'PortHandles');
        if portNum < 1 || portNum > numel(parentPortHandles.Inport)
            logMsg(logFile, 'WARN', 'Inport "%s" -> 상위 "%s" 포트 매핑 실패, 이름으로 대체', inportPath, parentPath);
            names = {get_param(inportHandle, 'Name')};
            constVals = {''};
            return;
        end

        [names, constVals] = traceSourceNames(parentPortHandles.Inport(portNum), swcPath, depth + 1, CFG, logFile);

        if isempty(names)
            % 상위로 올라가도 더 이상 못 찾으면(=상위에서 미연결) 이 Inport 이름을 그대로 사용
            names = {get_param(inportHandle, 'Name')};
            constVals = {''};
        end
    catch ME
        logMsg(logFile, 'ERROR', '서브시스템 경계 통과 중 오류 (%s): %s', inportPath, ME.message);
        names = {get_param(inportHandle, 'Name')};
        constVals = {''};
    end
end

function [names, constVals] = resolveFrom(fromHandle, fromPath, swcPath, depth, CFG, logFile)
    constVals = {};
    try
        gotoTag = get_param(fromHandle, 'GotoTag');
        model = bdroot(fromPath);
        % FollowLinks/LookUnderMasks 없이 찾으면 Goto가 마스킹된/라이브러리
        % 링크된 서브시스템 내부에 있을 때 아예 못 찾는다 (PhyVal 탐색 때와
        % 같은 이유) - 이게 "단순한 From/Goto인데도 인식을 못하는" 원인이었음.
        gotoBlocks = find_system(model, 'FollowLinks', 'on', 'LookUnderMasks', 'all', ...
            'BlockType', 'Goto', 'GotoTag', gotoTag);

        if isempty(gotoBlocks)
            logMsg(logFile, 'WARN', 'From "%s"(태그:%s)에 대응하는 Goto를 찾지 못함', fromPath, gotoTag);
            names = {get_param(fromHandle, 'Name')};
            constVals = {''};
            return;
        end

        gotoPath = gotoBlocks{1};
        if numel(gotoBlocks) > 1
            % 같은 태그 이름이 여러 서브시스템에서 로컬 스코프로 재사용된
            % 경우(흔함), 무작정 첫 번째가 아니라 From과 계층상 가장
            % 가까운(공통 상위 경로가 가장 긴) Goto를 고른다.
            gotoPath = selectClosestGoto(gotoBlocks, fromPath);
            logMsg(logFile, 'WARN', 'GotoTag "%s"에 대해 Goto 블록이 %d개 발견되어 가장 가까운 "%s"를 사용', ...
                gotoTag, numel(gotoBlocks), gotoPath);
        end

        % Goto가 swcPath 밖에 있으면(=scoped/global 태그로 SWC 외부에서 신호가
        % 들어오는 경우), 그 지점이 "SWC 단의 입력"이므로 더 거슬러 올라가지 않는다.
        insideSwc = strcmp(gotoPath, swcPath) || startsWith(gotoPath, [swcPath '/']);
        if ~insideSwc
            names = {get_param(gotoPath, 'Name')};
            constVals = {''};
            return;
        end

        gotoHandle = get_param(gotoPath, 'Handle');
        gotoPortHandles = get_param(gotoHandle, 'PortHandles');
        if isempty(gotoPortHandles.Inport)
            names = {get_param(gotoHandle, 'Name')};
            constVals = {''};
            return;
        end

        [names, constVals] = traceSourceNames(gotoPortHandles.Inport(1), swcPath, depth + 1, CFG, logFile);
        if isempty(names)
            names = {get_param(gotoHandle, 'Name')};
            constVals = {''};
        end
    catch ME
        logMsg(logFile, 'ERROR', 'From "%s" 처리 중 오류: %s', fromPath, ME.message);
        names = {sprintf('[오류:From처리실패:%s]', fromPath)};
        constVals = {''};
    end
end

function bestPath = selectClosestGoto(gotoBlocks, fromPath)
% 같은 GotoTag를 가진 Goto가 여러 개 있을 때(로컬 스코프 태그 재사용 등),
% From 블록과 계층 구조상 가장 가까운(경로 접두어가 가장 길게 겹치는) Goto를
% 고른다. Simulink의 local/scoped visibility 규칙을 완벽히 재현하지는
% 못하지만, 무작정 find_system이 찾은 첫 번째를 쓰는 것보다는 정확하다.
    fromTokens = strsplit(fromPath, '/');
    bestPath = gotoBlocks{1};
    bestScore = -1;
    for i = 1:numel(gotoBlocks)
        gotoTokens = strsplit(gotoBlocks{i}, '/');
        commonLen = 0;
        for k = 1:min(numel(fromTokens), numel(gotoTokens))
            if strcmp(fromTokens{k}, gotoTokens{k})
                commonLen = k;
            else
                break;
            end
        end
        if commonLen > bestScore
            bestScore = commonLen;
            bestPath = gotoBlocks{i};
        end
    end
end

function [names, constVals] = resolveGate(gateHandle, blockType, swcPath, depth, CFG, logFile)
    try
        gatePortHandles = get_param(gateHandle, 'PortHandles');
        collectedNames = {};
        collectedConsts = {};
        for p = 1:numel(gatePortHandles.Inport)
            [sub, subConst] = traceSourceNames(gatePortHandles.Inport(p), swcPath, depth + 1, CFG, logFile);
            if isempty(sub)
                sub = {sprintf('%s(미연결)', get_param(gateHandle, 'Name'))};
                subConst = {''};
            end
            collectedNames = [collectedNames, sub]; %#ok<AGROW>
            collectedConsts = [collectedConsts, subConst]; %#ok<AGROW>
        end
        if isempty(collectedNames)
            collectedNames = {get_param(gateHandle, 'Name')};
            collectedConsts = {''};
        end
        names = collectedNames;
        constVals = collectedConsts;
    catch ME
        logMsg(logFile, 'ERROR', '게이트 블록 "%s" 처리 중 오류: %s', get_param(gateHandle, 'Name'), ME.message);
        names = {'[오류:게이트처리실패]'};
        constVals = {''};
    end
end

% ====================================================================
% 로그 파일
% ====================================================================
function fid = openLogFile()
% 로그는 결과 엑셀 위치(outputFolder)와 무관하게, 항상 도구 폴더(이 파일이
% 들어있는 "data" 폴더의 한 단계 위)의 'logs' 하위 폴더에
% "log_<날짜시간>.txt" 이름으로 저장한다 (사용자가 항상 같은 위치/이름
% 규칙으로 찾을 수 있도록. data 폴더 안에 섞여 들어가지 않게 함).
    toolDir = fileparts(fileparts(mfilename('fullpath')));
    logsDir = fullfile(toolDir, 'logs');
    if ~exist(logsDir, 'dir')
        mkdir(logsDir);
    end

    ts = datestr(now, 'yyyymmdd_HHMMSS');
    logPath = fullfile(logsDir, sprintf('log_%s.txt', ts));

    fid = fopen(logPath, 'w', 'n', 'UTF-8');
    if fid == -1
        warning('Os_signal_extractor:LogOpenFailed', '로그 파일을 열지 못했습니다: %s', logPath);
        return;
    end
    fprintf('로그 파일: %s\n', logPath);
end

function logMsg(fid, level, fmt, varargin)
% 로그 파일에는 성공적으로 처리된 내용(mat/param 로드 완료, leaf 블록 개수,
% 게이트/경계 통과 등)은 남기지 않고, 특이사항(WARN/ERROR)과 실행 시작/종료
% 구분선(STAGE)만 남긴다. 콘솔에는 WARN/ERROR만 즉시 표시한다.
    ts = datestr(now, 'HH:MM:SS');
    line = sprintf('[%s][%s] %s', ts, level, sprintf(fmt, varargin{:}));

    if strcmp(level, 'WARN') || strcmp(level, 'ERROR')
        fprintf('%s\n', line); % 경고/에러는 콘솔에도 표시
    end
    if fid > 0 && (strcmp(level, 'WARN') || strcmp(level, 'ERROR') || strcmp(level, 'STAGE'))
        fprintf(fid, '%s\n', line);
    end
end

% ====================================================================
% 엑셀 출력
% ====================================================================
function outFile = writeCombinedExcel(resultsByModel, statsByModel, outputFolder, logFile, startTime)
% 선택된 모든 모델의 결과를 하나의 엑셀 파일에 "모델별 시트"로 나누어 저장하고,
% Summary 시트에 전체 실행 시간 + 모델별 통계를 함께 기록한다.
%
% 파일명은 항상 "result_<날짜시간>.xlsx" 로 고정한다 (모델 개수/이름과 무관).
    if ~exist(outputFolder, 'dir')
        mkdir(outputFolder);
    end

    modelNames = keys(resultsByModel);

    ts = datestr(now, 'yyyymmdd_HHMMSS');
    outFile = fullfile(outputFolder, sprintf('result_%s.xlsx', ts));
    usedSheetNames = {};
    for i = 1:numel(modelNames)
        modelName = modelNames{i};
        rows = resultsByModel(modelName);
        n = size(rows, 1);
        T = table((1:n)', rows(:, 1), rows(:, 2), rows(:, 3), rows(:, 4), rows(:, 5), ...
            'VariableNames', {'No', 'SubsystemPath', 'Input', 'ErrFlagInput', 'Output', 'Note'});

        sheetName = sanitizeSheetName(modelName, usedSheetNames);
        usedSheetNames{end + 1} = sheetName; %#ok<AGROW>
        writetable(T, outFile, 'Sheet', sheetName);
    end

    endTime = now;
    elapsedStr = datestr(endTime - startTime, 'HH:MM:SS');

    % Summary 시트 레이아웃 (고정 행 배치):
    %   A1  : 제목
    %   A3  : "실행 정보" 구분 표시
    %   A4  : 실행 정보 표 헤더(항목/값), A5:A7 데이터(3행 고정)
    %   A9  : "모델별 통계" 구분 표시
    %   A10 : 모델별 통계 표 헤더, A11~ 데이터
    writecell({'OS 신호 추출 결과 요약'}, outFile, 'Sheet', 'Summary', 'Range', 'A1');
    writecell({'실행 정보'}, outFile, 'Sheet', 'Summary', 'Range', 'A3');

    runInfoT = table( ...
        {'시작 시각'; '종료 시각'; '소요 시간'}, ...
        {datestr(startTime, 'yyyy-mm-dd HH:MM:SS'); datestr(endTime, 'yyyy-mm-dd HH:MM:SS'); elapsedStr}, ...
        'VariableNames', {'항목', '값'});
    writetable(runInfoT, outFile, 'Sheet', 'Summary', 'Range', 'A4');

    writecell({'모델별 통계'}, outFile, 'Sheet', 'Summary', 'Range', 'A9');

    modelStatRows = cell(numel(modelNames), 6);
    for i = 1:numel(modelNames)
        modelName = modelNames{i};
        rows = resultsByModel(modelName);
        stats = statsByModel(modelName);
        modelStatRows(i, :) = {modelName, size(rows, 1), stats.ground, stats.gateMulti, stats.dupRemoved, stats.unresolved};
    end
    statT = cell2table(modelStatRows, 'VariableNames', ...
        {'모델명', '전체건수', '그라운드', '다중입력', '중복제거', '미연결'});
    writetable(statT, outFile, 'Sheet', 'Summary', 'Range', 'A10');

    applyHeaderStyle(outFile, usedSheetNames, logFile);
    applySummaryStyle(outFile, logFile);

    logMsg(logFile, 'INFO', '엑셀 파일 생성 완료: %s', outFile);
    fprintf('엑셀 파일 생성 완료: %s\n', outFile);
end

function sheetName = sanitizeSheetName(name, usedNames)
% Excel 시트 이름 제약(31자 이내, 특정 특수문자 금지, 중복 불가)을 만족시킨다.
    nm = regexprep(name, '[:\\/?*\[\]]', '_');
    if isempty(nm)
        nm = 'Sheet';
    end
    if numel(nm) > 31
        nm = nm(1:31);
    end

    base = nm;
    suffix = 1;
    while any(strcmpi(usedNames, nm))
        suffixStr = sprintf('_%d', suffix);
        keepLen = max(1, 31 - numel(suffixStr));
        nm = [base(1:min(numel(base), keepLen)) suffixStr]; %#ok<AGROW>
        suffix = suffix + 1;
    end
    sheetName = nm;
end

function applySummaryStyle(filePath, logFile)
% Summary 시트 전용 서식: 제목(A1)은 크고 굵게, 구분 표시(A3, A9)는 굵게,
% 표 헤더(A4, A10 행)는 데이터 시트와 같은 파란 배경 + 굵게로 통일한다.
% (writeCombinedExcel의 고정 행 배치와 맞물려 있음 - 배치를 바꾸면 여기 행
%  번호도 같이 바꿔야 함)
    excelApp = [];
    workbook = [];
    try
        excelApp = actxserver('Excel.Application');
        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        workbook = excelApp.Workbooks.Open(filePath);

        headerColor = 217 + 225 * 256 + 242 * 65536; % #D9E1F2 연한 파랑 (COM Color는 BGR 순서)

        sheet = workbook.Sheets.Item('Summary');

        titleCell = sheet.Range('A1');
        titleCell.Font.Bold = true;
        titleCell.Font.Size = 14;

        sectionCells = {sheet.Range('A3'), sheet.Range('A9')};
        for i = 1:numel(sectionCells)
            sectionCells{i}.Font.Bold = true;
            sectionCells{i}.Font.Size = 11;
        end

        headerRows = {sheet.Range('A4:B4'), sheet.Range('A10:F10')};
        for i = 1:numel(headerRows)
            headerRows{i}.Interior.Color = headerColor;
            headerRows{i}.Font.Bold = true;
        end

        sheet.Columns.AutoFit;

        workbook.Save();
        workbook.Close(false);
        excelApp.Quit();
        delete(excelApp);
    catch ME
        logMsg(logFile, 'WARN', 'Summary 시트 서식 적용 실패 (%s): %s', filePath, ME.message);
        try
            if ~isempty(workbook)
                workbook.Close(false);
            end
            if ~isempty(excelApp)
                excelApp.Quit();
                delete(excelApp);
            end
        catch
        end
    end
end

function applyHeaderStyle(filePath, sheetNames, logFile)
% 헤더 행(1행)에 배경색 + 굵게를 적용해 본문과 시각적으로 구분되게 함.
% writetable만으로는 셀 서식을 지정할 수 없어 Excel COM으로 후처리한다.
% (Windows + Excel 설치 환경에서만 동작. 실패해도 결과 파일 자체는 이미
%  writetable로 생성되어 있으므로, 서식만 못 입히고 넘어간다.)
    excelApp = [];
    workbook = [];
    try
        excelApp = actxserver('Excel.Application');
        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        workbook = excelApp.Workbooks.Open(filePath);

        headerColor = 217 + 225 * 256 + 242 * 65536; % #D9E1F2 연한 파랑 (COM Color는 BGR 순서)

        for i = 1:numel(sheetNames)
            sheet = workbook.Sheets.Item(sheetNames{i});
            headerRow = sheet.Rows.Item(1);
            headerRow.Interior.Color = headerColor;
            headerRow.Font.Bold = true;
            sheet.Columns.AutoFit;
        end

        workbook.Save();
        workbook.Close(false);
        excelApp.Quit();
        delete(excelApp);
    catch ME
        logMsg(logFile, 'WARN', '엑셀 헤더 서식 적용 실패 (%s): %s', filePath, ME.message);
        try
            if ~isempty(workbook)
                workbook.Close(false);
            end
            if ~isempty(excelApp)
                excelApp.Quit();
                delete(excelApp);
            end
        catch
        end
    end
end
