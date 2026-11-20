function rootDir = make_example_project(rootDir)
%MAKE_EXAMPLE_PROJECT 추출기를 시험해 볼 수 있는 가짜 예제 프로젝트를 만든다.
%
%   rootDir = make_example_project()          % examples/example_project 에 생성
%   rootDir = make_example_project('C:\tmp')  % 원하는 위치에 생성
%
%   만들어지는 구조 (추출기가 기대하는 폴더 규칙과 같음)
%     example_project/
%       rte/demo_signals.mat          모델 파라미터가 참조하는 데이터
%       os/os_demo/os_demo.slx        예제 모델
%       os/os_demo/os_demo_param.m    모델 로드 전에 실행되는 파라미터 스크립트
%
%   예제 모델 os_demo 안의 Swc_Demo 서브시스템에는 추적 경우별 leaf 블록이 있다.
%     Out_VehSpd    : Gain 블록을 거쳐 SWC 입력까지 거슬러 올라감 + ErrFlag는 Goto/From
%     Out_DoorOpen  : OR 게이트 → 들어오는 두 신호를 모두 나열
%     Out_SelSpd    : Switch 게이트 → 세 입력을 모두 나열
%     Out_Reserved  : Ground → "ground" 표시
%     Out_DefaultLv : Constant → "constant(값)" 표시
%   이름에 _sch 가 들어간 Swc_Demo_sch 는 스케줄러로 보고 목록에서 빠져야 한다.

    if nargin < 1 || isempty(rootDir)
        rootDir = fullfile(fileparts(mfilename('fullpath')), 'example_project');
    end
    rteDir   = fullfile(rootDir, 'rte');
    modelDir = fullfile(rootDir, 'os', 'os_demo');
    if ~isfolder(rteDir),   mkdir(rteDir);   end
    if ~isfolder(modelDir), mkdir(modelDir); end

    % --- rte 데이터와 param 스크립트 ---
    DEMO_SPD_GAIN = 0.01;      %#ok<NASGU>
    DEMO_DEFAULT_LEVEL = 5;    %#ok<NASGU>
    save(fullfile(rteDir, 'demo_signals.mat'), 'DEMO_SPD_GAIN', 'DEMO_DEFAULT_LEVEL');

    fid = fopen(fullfile(modelDir, 'os_demo_param.m'), 'w', 'n', 'UTF-8');
    fprintf(fid, '%% os_demo 모델 파라미터 (rte/demo_signals.mat 로드 뒤에 실행됨)\n');
    fprintf(fid, 'os_demo_SpdGain = DEMO_SPD_GAIN;\n');
    fprintf(fid, 'os_demo_DefaultLevel = DEMO_DEFAULT_LEVEL;\n');
    fclose(fid);

    % --- 모델 ---
    mdl = 'os_demo';
    if bdIsLoaded(mdl), close_system(mdl, 0); end
    new_system(mdl);

    % 모델 최상위: 하네스 입력 → Swc_Demo
    harness = {'In_VehSpd_raw', 'In_SpdErr_raw', 'In_DoorFL_raw', 'In_DoorFR_raw', ...
               'In_SpdA_raw', 'In_SpdB_raw', 'In_SelCond_raw'};
    swc = [mdl '/Swc_Demo'];
    add_block('simulink/Ports & Subsystems/Subsystem', swc, 'Position', [300 40 460 400]);
    delete_line(swc, 'In1/1', 'Out1/1');
    delete_block([swc '/In1']);
    delete_block([swc '/Out1']);

    swcIn = {'VehSpd', 'SpdErr', 'DoorSts_FL', 'DoorSts_FR', 'SpdA', 'SpdB', 'SelCond'};
    for k = 1:numel(swcIn)
        add_block('simulink/Sources/In1', [swc '/' swcIn{k}], 'Position', [30 30+60*(k-1) 60 44+60*(k-1)]);
        add_block('simulink/Sources/In1', [mdl '/' harness{k}], 'Position', [100 50+50*(k-1) 130 64+50*(k-1)]);
    end

    % Out_VehSpd: PhyVal ← Gain ← VehSpd, ErrFlag ← From ← Goto ← SpdErr
    add_block('simulink/Math Operations/Gain', [swc '/SpdScale'], 'Gain', 'os_demo_SpdGain', 'Position', [150 25 190 50]);
    add_block('simulink/Signal Routing/Goto', [swc '/GotoSpdErr'], 'GotoTag', 'SpdErr', 'Position', [150 85 220 105]);
    add_block('simulink/Signal Routing/From', [swc '/FromSpdErr'], 'GotoTag', 'SpdErr', 'Position', [260 85 330 105]);
    addLeaf(swc, 'Out_VehSpd', true, [420 20 520 110]);
    add_line(swc, 'VehSpd/1', 'SpdScale/1');
    add_line(swc, 'SpdScale/1', 'Out_VehSpd/1');
    add_line(swc, 'SpdErr/1', 'GotoSpdErr/1');
    add_line(swc, 'FromSpdErr/1', 'Out_VehSpd/2');

    % Out_DoorOpen: PhyVal ← OR(DoorSts_FL, DoorSts_FR)
    add_block('simulink/Logic and Bit Operations/Logical Operator', [swc '/DoorOr'], ...
        'Operator', 'OR', 'Position', [200 160 240 220]);
    addLeaf(swc, 'Out_DoorOpen', false, [420 160 520 200]);
    add_line(swc, 'DoorSts_FL/1', 'DoorOr/1');
    add_line(swc, 'DoorSts_FR/1', 'DoorOr/2');
    add_line(swc, 'DoorOr/1', 'Out_DoorOpen/1');

    % Out_SelSpd: PhyVal ← Switch(SpdA, SelCond, SpdB)
    add_block('simulink/Signal Routing/Switch', [swc '/SpdSel'], 'Position', [200 260 240 340]);
    addLeaf(swc, 'Out_SelSpd', false, [420 280 520 320]);
    add_line(swc, 'SpdA/1', 'SpdSel/1');
    add_line(swc, 'SelCond/1', 'SpdSel/2');
    add_line(swc, 'SpdB/1', 'SpdSel/3');
    add_line(swc, 'SpdSel/1', 'Out_SelSpd/1');

    % Out_Reserved: PhyVal ← Ground
    add_block('simulink/Sources/Ground', [swc '/Gnd'], 'Position', [300 400 320 420]);
    addLeaf(swc, 'Out_Reserved', false, [420 390 520 430]);
    add_line(swc, 'Gnd/1', 'Out_Reserved/1');

    % Out_DefaultLv: PhyVal ← Constant
    add_block('simulink/Sources/Constant', [swc '/DefaultLv'], 'Value', 'os_demo_DefaultLevel', ...
        'Position', [280 460 330 480]);
    addLeaf(swc, 'Out_DefaultLv', false, [420 450 520 490]);
    add_line(swc, 'DefaultLv/1', 'Out_DefaultLv/1');

    % 최상위 하네스 연결
    for k = 1:numel(harness)
        add_line(mdl, [harness{k} '/1'], sprintf('Swc_Demo/%d', k));
    end

    % 스케줄러 서브시스템 (목록에서 빠져야 함)
    add_block('simulink/Ports & Subsystems/Subsystem', [mdl '/Swc_Demo_sch'], 'Position', [300 440 400 490]);

    save_system(mdl, fullfile(modelDir, [mdl '.slx']));
    close_system(mdl, 0);
    fprintf('예제 프로젝트를 만들었습니다: %s\n', rootDir);
end

function addLeaf(swc, name, withErrFlag, pos)
% 내부 Inport 이름이 PhyVal(그리고 ErrFlag)인 마스킹 서브시스템 = leaf 기능 블록
    blk = [swc '/' name];
    add_block('simulink/Ports & Subsystems/Subsystem', blk, 'Position', pos);
    set_param([blk '/In1'], 'Name', 'PhyVal');
    if withErrFlag
        add_block('simulink/Sources/In1', [blk '/ErrFlag'], 'Position', [30 100 60 114]);
        add_block('simulink/Sinks/Terminator', [blk '/TermErr'], 'Position', [120 100 140 120]);
        add_line(blk, 'ErrFlag/1', 'TermErr/1');
    end
    Simulink.Mask.create(blk);
end
